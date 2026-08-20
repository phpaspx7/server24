using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace MiniDbServer
{
    // ----------------------------------------------------------------------
    // Modello del record salvato nel database Access
    // ----------------------------------------------------------------------
    public class Record
    {
        public int Id { get; set; }
        public string Nome { get; set; } = "";
        public string Descrizione { get; set; } = "";
    }

    // Corpo della richiesta per l'endpoint di query SQL libera
    public class QueryRequest
    {
        public string Sql { get; set; } = "";
    }

    // Corpo della richiesta per l'endpoint di transazione: una sequenza di operazioni
    // SQL qualsiasi (SELECT, INSERT, UPDATE, DELETE, DDL...) da eseguire tutte insieme
    // o annullare tutte insieme.
    public class TransactionRequest
    {
        public List<string> Operazioni { get; set; } = new();
    }

    // Rappresenta una tabella del database con l'elenco dei suoi campi,
    // usata per mostrare nella pagina web la struttura del database connesso.
    public class TableSchema
    {
        public string Nome { get; set; } = "";
        public List<ColumnSchema> Colonne { get; set; } = new();
    }

    public class ColumnSchema
    {
        public string Nome { get; set; } = "";
        public string Tipo { get; set; } = "";
    }

    // Rappresenta una query salvata nel database Access (di selezione o di azione),
    // col relativo testo SQL letto tramite ADOX.
    public class QueryInfo
    {
        public string Nome { get; set; } = "";
        public string Tipo { get; set; } = ""; // "Selezione" (Views) oppure "Azione" (Procedures)
        public string Sql { get; set; } = "";
    }

    // Rappresenta un oggetto Access di cui, tramite MSysObjects, si può leggere
    // solo il nome (macro, report): il loro contenuto/design non è accessibile
    // dall'esterno di Access.
    public class ObjectInfo
    {
        public string Nome { get; set; } = "";
    }

    // Sollevata quando non è stato possibile ottenere l'accesso esclusivo al database
    // entro il tempo massimo di attesa, perché un'altra operazione è ancora in corso.
    // Distinta dalle altre eccezioni così gli endpoint possono restituire un errore
    // "occupato" (HTTP 503) invece di un generico errore 400/500.
    public class DbOccupatoException : Exception
    {
        public DbOccupatoException(string message) : base(message) { }
    }

    // ----------------------------------------------------------------------
    // Accesso al database Access (.accdb) tramite OleDb
    // ----------------------------------------------------------------------
    public class AccessDb
    {
        private readonly string _dbPath;
        private readonly string _connectionString;
        private const string TableName = "Records";

        // Timeout massimo per l'esecuzione di ogni comando SQL: 30 secondi.
        // Viene impostato esplicitamente su ogni comando creato, così una query lenta
        // fallisce dopo al più 30 secondi anziché restare bloccata a lungo.
        // Reso "internal" perché Program la usa per calcolare il timeout di sicurezza
        // (vedi Program.EseguiOperazioneDbAsync) leggermente più largo di questo.
        internal const int TimeoutSecondi = 30;

        // Percorso del file .accdb/.mdb a cui questa istanza è collegata.
        public string DbPath => _dbPath;

        public AccessDb(string dbPath)
        {
            _dbPath = dbPath;
            // "OLE DB Services=-4" disabilita il resource pooling nativo del provider
            // ACE OLEDB. Senza questo flag, OleDbConnection.ReleaseObjectPool() svuota
            // solo il pool gestito di .NET, ma il provider ACE mantiene comunque una
            // propria sessione nativa aperta sul file (a livello COM), che è la vera
            // causa per cui il .laccdb resta presente quando si cambia database.
            _connectionString = $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={_dbPath};OLE DB Services=-4;";
        }

        private OleDbConnection GetConnection() => new OleDbConnection(_connectionString);

        // Chiude davvero la connessione a questo database, svuotando il connection pool
        // di OleDb. Necessario perché ogni metodo (GetAll, Insert, ecc.) usa già
        // "using var conn" per aprire/chiudere la propria connessione, ma Dispose() su
        // una OleDbConnection NON chiude la connessione nativa: la restituisce solo al
        // pool, dove resta aperta "sotto" e continua a tenere il file .accdb bloccato
        // (file .laccdb). ReleaseObjectPool() forza la chiusura reale di tutte le
        // connessioni OleDb in pool nel processo, rilasciando il lock sul file.
        // Nota: a differenza di SqlConnection.ClearPool(), OleDb non permette di
        // svuotare il pool di una singola stringa di connessione: ReleaseObjectPool()
        // è un'operazione globale al processo, ma qui va benissimo perché il server
        // usa un solo database alla volta.
        // Va chiamato quando si cambia database o quando si chiude l'applicazione.
        public void ChiudiConnessioni()
        {
            try
            {
                OleDbConnection.ReleaseObjectPool();
            }
            catch
            {
                // Se fallisce non è un problema bloccante: il lock verrà comunque
                // rilasciato dal sistema operativo alla chiusura del processo.
            }

            // NOTA: qui NON va mai chiamato GC.Collect()/GC.WaitForPendingFinalizers().
            // GetQueries() ed EnsureDatabase() usano un oggetto COM (ADOX.Catalog): se
            // il thread di finalizzazione del GC deve marshalare la chiamata di rilascio
            // verso l'apartment COM originale e quell'apartment non è più raggiungibile,
            // WaitForPendingFinalizers() può bloccarsi per sempre, e da quel momento
            // QUALSIASI finalizzazione nel processo resta accodata dietro: il server
            // smette di rispondere in modo permanente dopo pochi cambi di database.
            // Il rilascio della connessione ADOX viene invece fatto in modo esplicito
            // e deterministico direttamente in GetQueries()/EnsureDatabase()
            // (catalog.ActiveConnection = null prima del FinalReleaseComObject),
            // senza bisogno di alcuna GC forzata qui.
        }

        // Crea un OleDbCommand già impostato con il timeout esteso di default.
        private static OleDbCommand CreaComando(string sql, OleDbConnection conn, OleDbTransaction? tx = null)
        {
            var cmd = new OleDbCommand(sql, conn, tx);
            cmd.CommandTimeout = TimeoutSecondi;
            return cmd;
        }

        // Crea il file .accdb e la tabella se non esistono già.
        // 'log' è opzionale: se fornito, riceve i messaggi di avanzamento
        // (usato per mostrarli nella finestra dell'applicazione).
        public void EnsureDatabase(Action<string>? log = null)
        {
            if (!File.Exists(_dbPath))
            {
                log?.Invoke($"Il file '{_dbPath}' non esiste: provo a crearlo...");
                try
                {
                    // Crea il file .accdb vuoto usando ADOX (richiede Access Database Engine installato)
                    Type? catalogType = Type.GetTypeFromProgID("ADOX.Catalog")
                        ?? throw new InvalidOperationException("ADOX non disponibile su questo sistema.");
                    dynamic catalog = Activator.CreateInstance(catalogType)!;
                    try
                    {
                        catalog.Create(_connectionString);
                        log?.Invoke("File .accdb creato correttamente.");
                    }
                    finally
                    {
                        // Catalog.Create() apre e mantiene una propria connessione nativa
                        // sul file appena creato: senza rilasciarla esplicitamente qui,
                        // il file resterebbe bloccato (.laccdb) finché il GC non decide
                        // di raccogliere l'oggetto COM, in un momento imprecisato.
                        try { catalog.ActiveConnection = null; } catch { /* ignorato */ }
                        if (Marshal.IsComObject(catalog)) Marshal.FinalReleaseComObject(catalog);
                    }
                }
                catch (Exception ex)
                {
                    throw new InvalidOperationException(
                        "Impossibile creare automaticamente il database Access. " +
                        "Crea manualmente un file .accdb vuoto con questo nome/percorso, oppure installa " +
                        "il 'Microsoft Access Database Engine 2016 Redistributable'. Dettaglio errore: " + ex.Message);
                }
            }

            // Nota: qui viene creato solo il file .accdb se mancante. Nessuna tabella
            // (né 'Records' né altre) viene mai creata automaticamente alla connessione:
            // il database collegato resta esattamente com'è, vuoto o con le tabelle
            // che contiene già.
        }

        // Legge un singolo Record dalla riga corrente del reader.
        // Usato sia da GetAll che da GetById per non duplicare la stessa mappatura.
        private static Record LeggiRecord(OleDbDataReader reader) => new()
        {
            Id = reader.GetInt32(0),
            Nome = reader.IsDBNull(1) ? "" : reader.GetString(1),
            Descrizione = reader.IsDBNull(2) ? "" : reader.GetString(2)
        };

        public List<Record> GetAll()
        {
            var list = new List<Record>();
            using var conn = GetConnection();
            conn.Open();
            using var cmd = CreaComando($"SELECT ID, Nome, Descrizione FROM {TableName} ORDER BY ID", conn);
            using var reader = cmd.ExecuteReader();
            while (reader.Read())
            {
                list.Add(LeggiRecord(reader));
            }
            return list;
        }

        public Record? GetById(int id)
        {
            using var conn = GetConnection();
            conn.Open();
            using var cmd = CreaComando($"SELECT ID, Nome, Descrizione FROM {TableName} WHERE ID = ?", conn);
            cmd.Parameters.AddWithValue("@id", id);
            using var reader = cmd.ExecuteReader();
            return reader.Read() ? LeggiRecord(reader) : null;
        }

        public int Insert(Record r)
        {
            using var conn = GetConnection();
            conn.Open();
            using (var cmd = CreaComando($"INSERT INTO {TableName} (Nome, Descrizione) VALUES (?, ?)", conn))
            {
                cmd.Parameters.AddWithValue("@nome", r.Nome ?? "");
                cmd.Parameters.AddWithValue("@descrizione", r.Descrizione ?? "");
                cmd.ExecuteNonQuery();
            }
            using (var idCmd = CreaComando("SELECT @@IDENTITY", conn))
            {
                var result = idCmd.ExecuteScalar();
                return Convert.ToInt32(result);
            }
        }

        public bool Update(int id, Record r)
        {
            using var conn = GetConnection();
            conn.Open();
            using var cmd = CreaComando($"UPDATE {TableName} SET Nome = ?, Descrizione = ? WHERE ID = ?", conn);
            cmd.Parameters.AddWithValue("@nome", r.Nome ?? "");
            cmd.Parameters.AddWithValue("@descrizione", r.Descrizione ?? "");
            cmd.Parameters.AddWithValue("@id", id);
            return cmd.ExecuteNonQuery() > 0;
        }

        public bool Delete(int id)
        {
            using var conn = GetConnection();
            conn.Open();
            using var cmd = CreaComando($"DELETE FROM {TableName} WHERE ID = ?", conn);
            cmd.Parameters.AddWithValue("@id", id);
            return cmd.ExecuteNonQuery() > 0;
        }

        // Esegue un comando già pronto: se il testo SQL è una SELECT restituisce le righe
        // trovate, altrimenti lo esegue come comando e restituisce il numero di righe
        // interessate. Usato sia per una singola query sia dentro una transazione,
        // evitando di duplicare la stessa logica in due punti diversi.
        private static object EseguiComando(OleDbCommand cmd, string sql)
        {
            bool isSelect = sql.TrimStart().StartsWith("SELECT", StringComparison.OrdinalIgnoreCase);
            if (!isSelect)
            {
                int righeInteressate = cmd.ExecuteNonQuery();
                return new { righeInteressate };
            }

            using var reader = cmd.ExecuteReader();
            var righe = new List<Dictionary<string, object?>>();
            while (reader.Read())
            {
                var riga = new Dictionary<string, object?>();
                for (int i = 0; i < reader.FieldCount; i++)
                {
                    riga[reader.GetName(i)] = reader.IsDBNull(i) ? null : reader.GetValue(i);
                }
                righe.Add(riga);
            }
            return righe;
        }

        // Esegue una qualsiasi query SQL scritta dall'utente:
        // - se è una SELECT, restituisce le righe trovate
        // - altrimenti (INSERT/UPDATE/DELETE/CREATE TABLE/ALTER/...) la esegue come comando
        //   e restituisce il numero di righe interessate.
        // Pensato per un uso locale/offline: non applica filtri di sicurezza sul testo SQL.
        public object ExecuteRawQuery(string sql)
        {
            using var conn = GetConnection();
            conn.Open();
            using var cmd = CreaComando(sql, conn);
            return EseguiComando(cmd, sql);
        }

        // Esegue una SEQUENZA di operazioni SQL (di qualsiasi tipo: SELECT, INSERT, UPDATE,
        // DELETE, CREATE TABLE, ecc.) dentro un'unica transazione:
        // - se TUTTE le operazioni vanno a buon fine, viene fatto un solo commit finale;
        // - se anche UNA sola operazione fallisce, tutte le altre già eseguite in questa
        //   chiamata vengono annullate (rollback) e il database torna com'era prima.
        // Questo evita che una sequenza (es. "svuota tabella" + tanti "inserisci riga")
        // possa interrompersi a metà lasciando i dati in uno stato incoerente.
        public List<object> EseguiTransazione(List<string> operazioni)
        {
            using var conn = GetConnection();
            conn.Open();
            using var tx = conn.BeginTransaction();

            var risultati = new List<object>();
            try
            {
                foreach (var sql in operazioni)
                {
                    if (string.IsNullOrWhiteSpace(sql)) continue;

                    using var cmd = CreaComando(sql, conn, tx);
                    risultati.Add(EseguiComando(cmd, sql));
                }

                tx.Commit();
                return risultati;
            }
            catch
            {
                // Anche solo un errore su una qualsiasi operazione annulla TUTTO
                // quello già eseguito in questa transazione.
                try { tx.Rollback(); } catch { /* connessione già chiusa/non valida */ }
                throw;
            }
        }

        // Restituisce l'elenco delle tabelle del database, ciascuna con il nome e il tipo
        // dei propri campi. Usato per mostrare la struttura del database nella pagina web
        // appena ci si connette. Le tabelle di sistema di Access (MSysXxx, ~xxx) vengono escluse.
        public List<TableSchema> GetSchema()
        {
            using var conn = GetConnection();
            conn.Open();

            var risultato = new List<TableSchema>();

            DataTable? tabelle = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object?[] { null, null, null, "TABLE" });
            if (tabelle == null) return risultato;

            foreach (DataRow rigaTabella in tabelle.Rows)
            {
                string nomeTabella = (string)rigaTabella["TABLE_NAME"];

                // Salta le tabelle di sistema/nascoste di Access
                if (nomeTabella.StartsWith("MSys", StringComparison.OrdinalIgnoreCase) || nomeTabella.StartsWith("~"))
                    continue;

                var tabella = new TableSchema { Nome = nomeTabella };

                DataTable? colonne = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Columns, new object?[] { null, null, nomeTabella, null });
                if (colonne != null)
                {
                    var righeOrdinate = colonne.Rows.Cast<DataRow>()
                        .OrderBy(r => Convert.ToInt32(r["ORDINAL_POSITION"]));

                    foreach (var rigaColonna in righeOrdinate)
                    {
                        tabella.Colonne.Add(new ColumnSchema
                        {
                            Nome = (string)rigaColonna["COLUMN_NAME"],
                            Tipo = DescriviTipo(rigaColonna)
                        });
                    }
                }

                risultato.Add(tabella);
            }

            return risultato;
        }

        // Legge l'elenco delle query salvate nel database (sia di selezione che di azione)
        // insieme al loro testo SQL, usando ADOX (lo stesso componente COM già usato per
        // creare il file .accdb in EnsureDatabase). Richiede ADOX/Access Database Engine
        // installato: se non disponibile, lancia un'eccezione che il chiamante può gestire
        // senza far fallire il resto della risposta (tabelle/macro/report).
        public List<QueryInfo> GetQueries()
        {
            Type? catalogType = Type.GetTypeFromProgID("ADOX.Catalog")
                ?? throw new InvalidOperationException("ADOX non disponibile su questo sistema: impossibile leggere le query.");

            dynamic catalog = Activator.CreateInstance(catalogType)!;
            var risultato = new List<QueryInfo>();
            try
            {
                catalog.ActiveConnection = _connectionString;

                // Query di selezione (SELECT) -> collezione Views
                foreach (dynamic view in catalog.Views)
                {
                    string nome = (string)view.Name;
                    if (nome.StartsWith("~", StringComparison.Ordinal)) continue; // query/viste temporanee interne
                    string sql = "";
                    try { sql = (string)view.Command.CommandText; } catch { /* testo SQL non sempre esposto */ }
                    risultato.Add(new QueryInfo { Nome = nome, Tipo = "Selezione", Sql = sql });
                }

                // Query di azione (INSERT/UPDATE/DELETE/CREATE TABLE/UNION...) -> collezione Procedures
                foreach (dynamic proc in catalog.Procedures)
                {
                    string nome = (string)proc.Name;
                    if (nome.StartsWith("~", StringComparison.Ordinal)) continue;
                    string sql = "";
                    try { sql = (string)proc.Command.CommandText; } catch { /* testo SQL non sempre esposto */ }
                    risultato.Add(new QueryInfo { Nome = nome, Tipo = "Azione", Sql = sql });
                }
            }
            finally
            {
                // Prima di rilasciare l'oggetto COM, si stacca esplicitamente la
                // connessione (ActiveConnection = null): questo chiude subito e in
                // modo deterministico la sessione nativa ADO sul file .accdb, invece
                // di lasciare che sia il garbage collector a farlo in un momento
                // imprecisato (comportamento che, tra l'altro, può causare blocchi:
                // vedi il commento in ChiudiConnessioni()).
                try { catalog.ActiveConnection = null; } catch { /* ignorato */ }
                if (Marshal.IsComObject(catalog)) Marshal.FinalReleaseComObject(catalog);
            }

            return risultato.OrderBy(q => q.Nome, StringComparer.OrdinalIgnoreCase).ToList();
        }

        // Legge dalla tabella di sistema MSysObjects il solo NOME degli oggetti del tipo
        // indicato (macro o report). Il contenuto/design di questi oggetti non è leggibile
        // dall'esterno di Access: solo l'elenco dei nomi presenti nel database.
        //
        // MSysObjects è protetta di default: prima della SELECT viene tentato un
        // "GRANT SELECT ON MSysObjects TO Admin", trucco noto che sblocca la lettura per
        // la connessione corrente. Non è permanente: va ripetuto a ogni connessione.
        // Se anche questo fallisse, la SELECT successiva solleverà l'eccezione originale
        // di permesso negato, che il chiamante può mostrare come messaggio esplicativo.
        private List<ObjectInfo> LeggiOggettiMSysObjects(int tipo)
        {
            using var conn = GetConnection();
            conn.Open();

            try
            {
                using var grantCmd = CreaComando("GRANT SELECT ON MSysObjects TO Admin", conn);
                grantCmd.ExecuteNonQuery();
            }
            catch { /* ignorato: si tenta comunque la lettura sotto */ }

            var risultato = new List<ObjectInfo>();
            using var cmd = CreaComando(
                "SELECT Name FROM MSysObjects WHERE Type = ? AND Left(Name,1) <> '~' AND Left(Name,4) <> 'MSys' ORDER BY Name",
                conn);
            cmd.Parameters.AddWithValue("@tipo", tipo);
            using var reader = cmd.ExecuteReader();
            while (reader.Read())
            {
                risultato.Add(new ObjectInfo { Nome = reader.IsDBNull(0) ? "" : reader.GetString(0) });
            }
            return risultato;
        }

        // Tipo -32766 in MSysObjects = Macro
        public List<ObjectInfo> GetMacro() => LeggiOggettiMSysObjects(-32766);

        // Tipo -32764 in MSysObjects = Report
        public List<ObjectInfo> GetReport() => LeggiOggettiMSysObjects(-32764);

        // Traduce il codice numerico del tipo di dato OleDb in un nome leggibile in italiano
        // (es. "Testo (255)", "Numero intero", "Data/Ora", ...).
        private static string DescriviTipo(DataRow rigaColonna)
        {
            var tipoOleDb = (OleDbType)Convert.ToInt32(rigaColonna["DATA_TYPE"]);

            string nome = tipoOleDb switch
            {
                OleDbType.WChar or OleDbType.VarWChar or OleDbType.Char or OleDbType.VarChar => "Testo",
                OleDbType.LongVarWChar or OleDbType.LongVarChar => "Memo",
                OleDbType.Integer => "Numero intero",
                OleDbType.SmallInt or OleDbType.UnsignedSmallInt => "Intero breve",
                OleDbType.TinyInt or OleDbType.UnsignedTinyInt => "Byte",
                OleDbType.Double or OleDbType.Single or OleDbType.Decimal or OleDbType.Numeric => "Numero decimale",
                OleDbType.Currency => "Valuta",
                OleDbType.Boolean => "Sì/No",
                OleDbType.Date or OleDbType.DBTimeStamp or OleDbType.DBDate => "Data/Ora",
                OleDbType.LongVarBinary or OleDbType.VarBinary or OleDbType.Binary => "Allegato/Binario",
                OleDbType.Guid => "GUID",
                _ => tipoOleDb.ToString()
            };

            bool haLunghezzaTesto = tipoOleDb is OleDbType.WChar or OleDbType.VarWChar or OleDbType.Char or OleDbType.VarChar;
            if (haLunghezzaTesto && rigaColonna["CHARACTER_MAXIMUM_LENGTH"] != DBNull.Value)
            {
                nome += $" ({Convert.ToInt32(rigaColonna["CHARACTER_MAXIMUM_LENGTH"])})";
            }

            return nome;
        }
    }

    // Corpo della richiesta per selezionare quale database .accdb/.mdb usare
    public class DatabaseSelectRequest
    {
        public string Percorso { get; set; } = "";
        public bool Crea { get; set; } = false; // se true, crea il file .accdb vuoto (senza tabelle) se non esiste
    }

    // ----------------------------------------------------------------------
    // Piccola finestra dell'applicazione: mostra il log del server e un pulsante
    // per aprire la pagina web nel browser predefinito.
    // ----------------------------------------------------------------------
    public class MainForm : Form
    {
        private readonly TextBox _logTextBox;
        private readonly Button _btnApriBrowser;

        public MainForm()
        {
            Text = "MiniDbServer";
            Width = 560;
            Height = 400;
            MinimumSize = new Size(400, 250);
            StartPosition = FormStartPosition.CenterScreen;

            _logTextBox = new TextBox
            {
                Multiline = true,
                ReadOnly = true,
                ScrollBars = ScrollBars.Vertical,
                Dock = DockStyle.Fill,
                Font = new Font("Consolas", 9F),
                BackColor = Color.Black,
                ForeColor = Color.FromArgb(140, 255, 140),
                BorderStyle = BorderStyle.FixedSingle
            };

            _btnApriBrowser = new Button
            {
                Text = "Apri nel browser",
                Dock = DockStyle.Bottom,
                Height = 42,
                Font = new Font("Segoe UI", 10F, FontStyle.Bold)
            };
            _btnApriBrowser.Click += (s, e) => ApriBrowser();

            Controls.Add(_logTextBox);
            Controls.Add(_btnApriBrowser);
        }

        private void ApriBrowser()
        {
            try
            {
                Process.Start(new ProcessStartInfo
                {
                    FileName = Program.BaseUrl + "home.htm",
                    UseShellExecute = true
                });
            }
            catch (Exception ex)
            {
                AggiungiLog("Impossibile aprire il browser: " + ex.Message);
            }
        }

        // Aggiunge una riga al log. Sicuro da chiamare anche da thread diversi
        // da quello dell'interfaccia grafica (es. dal thread del server web).
        public void AggiungiLog(string messaggio)
        {
            string riga = $"[{DateTime.Now:HH:mm:ss}] {messaggio}{Environment.NewLine}";

            if (_logTextBox.IsHandleCreated && _logTextBox.InvokeRequired)
            {
                try { _logTextBox.Invoke(new Action(() => ScriviRiga(riga))); }
                catch (ObjectDisposedException) { /* finestra già chiusa */ }
                catch (InvalidOperationException) { /* handle non ancora pronto/finestra in chiusura */ }
            }
            else
            {
                ScriviRiga(riga);
            }
        }

        private void ScriviRiga(string riga)
        {
            _logTextBox.AppendText(riga);
        }
    }

    // ----------------------------------------------------------------------
    // Server HTTP incorporato
    // ----------------------------------------------------------------------
    public class Program
    {
        // Indirizzo su cui il server ascolta e su cui si apre il browser.
        // 127.0.0.1 (loopback) invece di "localhost": evita risoluzioni DNS
        // inutili ed è sempre lo stesso indirizzo, indipendentemente dalla
        // configurazione di rete del PC.
        public const string BaseUrl = "http://127.0.0.1:5000/";

        private static AccessDb? _db;
        private static string _exeDir = "";
        private static readonly object _dbLock = new();

        // Il motore Jet/ACE (Microsoft.ACE.OLEDB.12.0) usato per i file .accdb/.mdb NON è
        // thread-safe: se due richieste HTTP arrivano quasi insieme (HandleRequestAsync
        // gira ogni richiesta in background con "_ = HandleRequestAsync(...)", quindi
        // possono sovrapporsi) e ciascuna esegue un comando OleDb in parallelo, il
        // provider nativo può bloccarsi internamente in modo permanente. Da quel momento
        // anche tutte le richieste successive restano in coda dietro quella bloccata: il
        // server sembra "non rispondere più a nessuna query", pur restando in esecuzione.
        // Questo semaforo (capacità 1) serializza TUTTE le operazioni che toccano
        // effettivamente il database (CRUD, query libere, transazioni, lettura schema,
        // cambio database), così un solo comando OleDb per volta arriva al motore Jet/ACE.
        private static readonly SemaphoreSlim _dbSemaphore = new(1, 1);
        private static MainForm? _mainForm;
        private static HttpListener? _listener;
        private static readonly CancellationTokenSource _cts = new();

        // Client HTTP condiviso e riutilizzato per le richieste verso pagine esterne
        // (endpoint /apri?url=...). Va creato una sola volta come campo statico: aprirne
        // uno nuovo a ogni richiesta esaurirebbe le socket disponibili sotto carico.
        private static readonly HttpClient _httpClient = new HttpClient { Timeout = TimeSpan.FromSeconds(30) };

        [STAThread]
        public static void Main(string[] args)
        {
            ApplicationConfiguration.Initialize();

            _mainForm = new MainForm();

            // Quando la finestra viene chiusa, ferma subito il server (listener HTTP
            // e ciclo di ascolto), così non restano connessioni o richieste in sospeso.
            _mainForm.FormClosing += (s, e) => ArrestaServer();

            // Il server web gira su un thread separato, così non blocca la finestra grafica.
            _ = Task.Run(() => AvviaServerAsync(args));

            Application.Run(_mainForm);

            // La finestra è chiusa: assicurati che tutto sia fermato (listener, ciclo di
            // ascolto, eventuali connessioni) e chiudi davvero il processo, senza lasciare
            // in vita thread di sfondo o la porta occupata.
            ArrestaServer();
            Environment.Exit(0);
        }

        // Ferma in modo pulito il server: chiude il listener HTTP (interrompendo eventuali
        // GetContextAsync() in attesa) e segnala al ciclo di ascolto di uscire.
        // Sicura da chiamare più volte.
        private static void ArrestaServer()
        {
            try
            {
                if (!_cts.IsCancellationRequested)
                    _cts.Cancel();
            }
            catch { /* ignora */ }

            try
            {
                if (_listener != null && _listener.IsListening)
                    _listener.Stop();
                _listener?.Close();
            }
            catch { /* il listener potrebbe essere già stato chiuso */ }

            // Chiude davvero la connessione al database attualmente in uso (svuota il
            // pool OleDb), così il file .accdb non resta bloccato (.laccdb residuo)
            // dopo la chiusura del programma.
            try
            {
                lock (_dbLock)
                {
                    _db?.ChiudiConnessioni();
                }
            }
            catch { /* ignorato: siamo comunque in fase di chiusura del processo */ }
        }

        // Cerca nella cartella dell'eseguibile tutti i file .accdb/.mdb, in ordine di nome
        // (prima gli .accdb poi gli .mdb). Non crea né modifica nulla: si limita a leggere
        // l'elenco dei file già presenti.
        private static IEnumerable<string> TrovaDatabaseNellaCartella()
        {
            return Directory.GetFiles(_exeDir, "*.accdb")
                .Concat(Directory.GetFiles(_exeDir, "*.mdb"))
                .OrderBy(f => f);
        }

        // Collega il server al file .accdb/.mdb indicato.
        // - Se il file non esiste e 'creaSeMancante' è false, l'operazione fallisce
        //   (nessun file viene creato "ex novo").
        // - Se il file non esiste e 'creaSeMancante' è true, viene creato un nuovo file
        //   .accdb vuoto (senza alcuna tabella) prima di collegarsi.
        // - Se il file esiste già, non viene toccato in alcun modo: nessuna tabella
        //   viene creata o modificata al momento della connessione.
        private static void ConnettiA(string percorso, bool creaSeMancante)
        {
            if (!File.Exists(percorso) && !creaSeMancante)
                throw new FileNotFoundException($"Il file '{percorso}' non esiste.");

            var nuovoDb = new AccessDb(percorso);
            nuovoDb.EnsureDatabase(Log);

            lock (_dbLock)
            {
                // Prima di agganciarsi al nuovo database, chiude davvero la connessione
                // al database precedentemente in uso (svuota il pool), così il file
                // .accdb di prima resta libero (nessun .laccdb residuo) subito dopo il cambio.
                _db?.ChiudiConnessioni();
                _db = nuovoDb;
            }
        }

        // Scrive un messaggio solo nella casella di log della finestra grafica
        // (nessuna console: l'app è WinExe e non ne apre una).
        private static void Log(string messaggio)
        {
            _mainForm?.AggiungiLog(messaggio);
        }

        // Timeout di sicurezza per ogni operazione sul database: un po' più largo del
        // CommandTimeout impostato su ogni OleDbCommand (AccessDb.TimeoutSecondi = 30s).
        // Serve solo come ultima rete: nel caso normale, un comando che va per le lunghe
        // fallisce da solo con un OleDbException di timeout ben prima di questo limite.
        private static readonly TimeSpan TimeoutDiSicurezza = TimeSpan.FromSeconds(AccessDb.TimeoutSecondi + 5);

        // Tempo massimo di attesa per ottenere il semaforo del database prima di
        // rinunciare e segnalare al chiamante che il database è occupato da un'altra
        // operazione, invece di lasciarlo in attesa indefinitamente in coda.
        private static readonly TimeSpan TimeoutAttesaSemaforo = TimeSpan.FromSeconds(10);

        // Punto centrale da cui passa OGNI accesso al database. Fa due cose:
        // 1) Serializza le operazioni con _dbSemaphore, perché il motore Jet/ACE non
        //    tollera comandi concorrenti sullo stesso file.
        // 2) Esegue l'operazione su un thread separato con un timeout di sicurezza: se
        //    una query (tipicamente un'azione andata storta) blocca il comando nativo
        //    senza che nemmeno il CommandTimeout la faccia fallire, questo è l'unico
        //    modo per non restare bloccati per sempre. Non è possibile interrompere in
        //    modo sicuro una chiamata OleDb nativa bloccata (in .NET moderno non esiste
        //    più un equivalente affidabile di Thread.Abort): quindi, se scatta il
        //    timeout, si rilascia SUBITO il semaforo (le richieste successive possono
        //    procedere) e si segnala un errore al chiamante, mentre il comando rimasto
        //    bloccato continua per conto suo in background e il suo risultato, quando
        //    (e se) arriverà, viene semplicemente ignorato.
        private static async Task<T> EseguiOperazioneDbAsync<T>(Func<T> operazione)
        {
            // Se il database è già occupato da un'altra operazione, si aspetta al
            // massimo TimeoutAttesaSemaforo (10s): oltre, si rinuncia e si segnala
            // subito "occupato" al chiamante, senza restare in coda indefinitamente.
            bool haOttenutoSemaforo = await _dbSemaphore.WaitAsync(TimeoutAttesaSemaforo);
            if (!haOttenutoSemaforo)
            {
                throw new DbOccupatoException(
                    $"Database occupato: un'altra operazione è ancora in corso da più di {TimeoutAttesaSemaforo.TotalSeconds:0} secondi. Riprova tra qualche istante.");
            }

            bool semaforoGiaRilasciato = false;
            try
            {
                Task<T> dbTask = Task.Run(operazione);
                Task completata = await Task.WhenAny(dbTask, Task.Delay(TimeoutDiSicurezza));

                if (completata != dbTask)
                {
                    _dbSemaphore.Release();
                    semaforoGiaRilasciato = true;

                    // Fire-and-forget: quando il comando bloccato terminerà (se mai
                    // succederà), si limita a loggare l'esito, senza toccare più il
                    // semaforo (evita un doppio rilascio).
                    _ = dbTask.ContinueWith(t =>
                    {
                        if (t.IsFaulted)
                            Log("Un comando rimasto bloccato oltre il timeout di sicurezza è terminato in ritardo con errore: " + t.Exception?.GetBaseException().Message);
                        else
                            Log("Un comando rimasto bloccato oltre il timeout di sicurezza è terminato in ritardo: risultato ignorato.");
                    }, TaskScheduler.Default);

                    throw new TimeoutException(
                        $"Il comando non ha risposto entro {TimeoutDiSicurezza.TotalSeconds:0} secondi ed è stato considerato bloccato " +
                        "(tipico di una query di azione andata storta che manda in stallo il motore Access). " +
                        "Le query successive possono comunque procedere; se il problema persiste, chiudi e riapri MiniDbServer per liberare il file.");
                }

                return await dbTask; // propaga il risultato, o l'eccezione originale della query
            }
            finally
            {
                if (!semaforoGiaRilasciato) _dbSemaphore.Release();
            }
        }

        private static async Task AvviaServerAsync(string[] args)
        {
            _exeDir = AppDomain.CurrentDomain.BaseDirectory;

            Log("=== MiniDbServer ===");

            // All'avvio NON viene mai creato un nuovo database: si cercano solo file
            // .accdb/.mdb già presenti nella cartella dell'eseguibile. Se ce n'è almeno
            // uno, ci si connette automaticamente al primo; altrimenti si resta senza
            // database collegato, in attesa che l'utente ne scelga (o crei) uno dalla pagina web.
            string? primoTrovato = TrovaDatabaseNellaCartella().FirstOrDefault();
            if (primoTrovato != null)
            {
                try
                {
                    ConnettiA(primoTrovato, creaSeMancante: false);
                    Log($"Connesso automaticamente a: {primoTrovato}");
                }
                catch (Exception ex)
                {
                    Log("ERRORE nella connessione automatica al database trovato:");
                    Log(ex.Message);
                }
            }
            else
            {
                Log($"Nessun file .accdb/.mdb trovato in '{_exeDir}'. Nessun database collegato: selezionane uno dalla pagina web.");
            }

            string prefix = Program.BaseUrl;
            _listener = new HttpListener();
            _listener.Prefixes.Add(prefix);

            try
            {
                _listener.Start();
            }
            catch (HttpListenerException ex)
            {
                Log("Impossibile avviare il server: " + ex.Message);
                Log("Prova ad avviare l'exe come amministratore, oppure verifica che la porta 5000 sia libera.");
                return;
            }

            Log($"Server avviato su {prefix}");
            Log($"Usa il pulsante 'Apri nel browser' per vedere la pagina, oppure vai su {Program.BaseUrl}home.htm");

            while (!_cts.IsCancellationRequested)
            {
                try
                {
                    var context = await _listener.GetContextAsync();
                    _ = HandleRequestAsync(context); // gestisce ogni richiesta in background
                }
                catch (Exception) when (_cts.IsCancellationRequested)
                {
                    // Chiusura in corso (finestra chiusa): esce dal ciclo senza segnalare errori.
                    break;
                }
                catch (HttpListenerException)
                {
                    // Il listener è stato fermato/chiuso: esce dal ciclo.
                    break;
                }
                catch (ObjectDisposedException)
                {
                    // Il listener è stato chiuso durante l'attesa: esce dal ciclo.
                    break;
                }
                catch (Exception ex)
                {
                    // Qualunque altro errore imprevisto su GetContextAsync() (es. un
                    // problema di rete transitorio) NON deve far morire il ciclo di
                    // ascolto: prima di questo catch, un'eccezione di questo tipo
                    // usciva dal while, la Task lanciata con Task.Run() falliva in
                    // modo silenzioso (nessuno la stava "await"-ando) e il server
                    // smetteva di rispondere a qualunque richiesta, dando l'impressione
                    // che "cadesse la linea" dopo un po' di tempo, senza però chiudere
                    // la finestra dell'applicazione. Qui invece l'errore viene solo
                    // loggato e il ciclo continua ad accettare nuove richieste.
                    Log("Errore nel ciclo di ascolto (ignorato, il server continua): " + ex.Message);
                }
            }

            Log("Server fermato.");
        }

        private static async Task HandleRequestAsync(HttpListenerContext context)
        {
            var req = context.Request;
            var res = context.Response;

            try
            {
                string path = req.Url?.AbsolutePath ?? "/";
                string method = req.HttpMethod;

                Log($"{method} {path}");

                res.Headers.Add("Access-Control-Allow-Origin", "*");

                if (method == "GET" && (path == "/" || path == "/index.html" || path == "/home.htm"))
                {
                    await ServiHomeHtmAsync(res);
                    return;
                }

                if (path.StartsWith("/api/records"))
                {
                    await HandleApiAsync(context, path, method);
                    return;
                }

                if (path == "/api/query" && method == "POST")
                {
                    await HandleQueryAsync(context);
                    return;
                }

                if (path == "/api/transazione" && method == "POST")
                {
                    await HandleTransactionAsync(context);
                    return;
                }

                if (path == "/api/database/elenco" && method == "GET")
                {
                    await HandleListDatabasesAsync(context);
                    return;
                }

                if (path == "/api/database/attuale" && method == "GET")
                {
                    await WriteJsonAsync(res, new { connesso = _db != null, percorso = _db?.DbPath });
                    return;
                }

                if (path == "/api/database/seleziona" && method == "POST")
                {
                    await HandleSelectDatabaseAsync(context);
                    return;
                }

                if (path == "/api/database/schema" && method == "GET")
                {
                    await HandleSchemaAsync(context);
                    return;
                }

                // Apre una pagina/file che NON si trova nella cartella dell'eseguibile:
                // - "?percorso=" per un file ovunque sul disco (o su una condivisione di rete)
                // - "?url="      per una pagina esterna su Internet (il server fa da proxy)
                if (path == "/apri" && method == "GET")
                {
                    await HandleApriAsync(context);
                    return;
                }

                // Qualsiasi altra richiesta GET viene cercata come file statico nella
                // cartella dell'eseguibile (o in una sua sottocartella): permette di
                // aprire pagine web aggiuntive (es. carica_versato.htm) senza dover
                // aggiungere una rotta dedicata per ciascuna.
                if (method == "GET" && await ServiFileStaticoAsync(res, path))
                {
                    return;
                }

                res.StatusCode = 404;
                await WriteJsonAsync(res, new { errore = "Non trovato" });
            }
            catch (DbOccupatoException ex)
            {
                Log("OCCUPATO: " + ex.Message);
                try
                {
                    res.StatusCode = 503; // Service Unavailable: riprovare più tardi
                    await WriteJsonAsync(res, new { successo = false, occupato = true, errore = ex.Message });
                }
                catch { /* connessione già chiusa */ }
            }
            catch (Exception ex)
            {
                Log("ERRORE: " + ex.Message);
                try
                {
                    res.StatusCode = 500;
                    await WriteJsonAsync(res, new { errore = ex.Message });
                }
                catch { /* connessione già chiusa */ }
            }
            finally
            {
                res.OutputStream.Close();
            }
        }

        private static async Task HandleApiAsync(HttpListenerContext context, string path, string method)
        {
            var req = context.Request;
            var res = context.Response;

            var db = _db;
            if (db == null)
            {
                res.StatusCode = 400;
                await WriteJsonAsync(res, new { errore = "Nessun database connesso. Selezionane uno dalla pagina principale." });
                return;
            }

            // /api/records         -> lista (GET) / creazione (POST)
            // /api/records/{id}    -> singolo (GET) / modifica (PUT) / cancellazione (DELETE)
            string[] parts = path.Trim('/').Split('/');
            int? id = null;
            if (parts.Length == 3 && int.TryParse(parts[2], out int parsedId)) id = parsedId;

            // Corpo della richiesta (POST/PUT) letto PRIMA di prendere il semaforo: legge solo
            // dallo stream HTTP, non tocca il database, quindi non ha senso tenere bloccate le
            // altre richieste nel frattempo.
            Record? newRec = null, updated = null;
            if (method == "POST")
            {
                newRec = await ReadJsonAsync<Record>(req);
                if (newRec == null) { res.StatusCode = 400; await WriteJsonAsync(res, new { errore = "Corpo non valido" }); return; }
            }
            else if (method == "PUT" && id != null)
            {
                updated = await ReadJsonAsync<Record>(req);
                if (updated == null) { res.StatusCode = 400; await WriteJsonAsync(res, new { errore = "Corpo non valido" }); return; }
            }

            // Il motore Jet/ACE non tollera comandi concorrenti sullo stesso file: ogni
            // chiamata ad AccessDb passa da EseguiOperazioneDbAsync, che serializza le
            // operazioni e protegge da un eventuale comando rimasto bloccato.
            switch (method)
            {
                case "GET" when id == null:
                    var lista = await EseguiOperazioneDbAsync(() => db.GetAll());
                    await WriteJsonAsync(res, lista);
                    break;

                case "GET" when id != null:
                    var rec = await EseguiOperazioneDbAsync(() => db.GetById(id.Value));
                    if (rec == null) { res.StatusCode = 404; await WriteJsonAsync(res, new { errore = "Record non trovato" }); }
                    else await WriteJsonAsync(res, rec);
                    break;

                case "POST":
                    int newId = await EseguiOperazioneDbAsync(() => db.Insert(newRec!));
                    newRec!.Id = newId;
                    res.StatusCode = 201;
                    await WriteJsonAsync(res, newRec);
                    break;

                case "PUT" when id != null:
                    bool ok = await EseguiOperazioneDbAsync(() => db.Update(id.Value, updated!));
                    if (!ok) { res.StatusCode = 404; await WriteJsonAsync(res, new { errore = "Record non trovato" }); }
                    else { updated!.Id = id.Value; await WriteJsonAsync(res, updated); }
                    break;

                case "DELETE" when id != null:
                    bool deleted = await EseguiOperazioneDbAsync(() => db.Delete(id.Value));
                    if (!deleted) { res.StatusCode = 404; await WriteJsonAsync(res, new { errore = "Record non trovato" }); }
                    else await WriteJsonAsync(res, new { successo = true });
                    break;

                default:
                    res.StatusCode = 405;
                    await WriteJsonAsync(res, new { errore = "Metodo non supportato" });
                    break;
            }
        }

        private static async Task HandleQueryAsync(HttpListenerContext context)
        {
            var req = context.Request;
            var res = context.Response;

            var db = _db;
            if (db == null)
            {
                res.StatusCode = 400;
                await WriteJsonAsync(res, new { successo = false, errore = "Nessun database connesso. Selezionane uno dalla pagina principale." });
                return;
            }

            var body = await ReadJsonAsync<QueryRequest>(req);
            if (body == null || string.IsNullOrWhiteSpace(body.Sql))
            {
                res.StatusCode = 400;
                await WriteJsonAsync(res, new { successo = false, errore = "Devi specificare il campo 'sql'." });
                return;
            }

            try
            {
                var risultato = await EseguiOperazioneDbAsync(() => db.ExecuteRawQuery(body.Sql));
                await WriteJsonAsync(res, new { successo = true, risultato });
            }
            catch (DbOccupatoException ex)
            {
                res.StatusCode = 503; // Service Unavailable: riprovare più tardi
                await WriteJsonAsync(res, new { successo = false, occupato = true, errore = ex.Message });
            }
            catch (Exception ex)
            {
                res.StatusCode = 400;
                await WriteJsonAsync(res, new { successo = false, errore = ex.Message });
            }
        }

        private static async Task HandleTransactionAsync(HttpListenerContext context)
        {
            var req = context.Request;
            var res = context.Response;

            var db = _db;
            if (db == null)
            {
                res.StatusCode = 400;
                await WriteJsonAsync(res, new { successo = false, errore = "Nessun database connesso. Selezionane uno dalla pagina principale." });
                return;
            }

            var body = await ReadJsonAsync<TransactionRequest>(req);
            if (body == null || body.Operazioni == null || body.Operazioni.Count == 0)
            {
                res.StatusCode = 400;
                await WriteJsonAsync(res, new { successo = false, errore = "Devi specificare almeno un'operazione nel campo 'operazioni'." });
                return;
            }

            try
            {
                var risultati = await EseguiOperazioneDbAsync(() => db.EseguiTransazione(body.Operazioni));
                await WriteJsonAsync(res, new { successo = true, operazioniEseguite = risultati.Count, risultati });
            }
            catch (DbOccupatoException ex)
            {
                res.StatusCode = 503; // Service Unavailable: riprovare più tardi
                await WriteJsonAsync(res, new { successo = false, occupato = true, errore = ex.Message });
            }
            catch (Exception ex)
            {
                // Nessuna delle operazioni è stata mantenuta: rollback già avvenuto in EseguiTransazione.
                res.StatusCode = 400;
                await WriteJsonAsync(res, new { successo = false, errore = ex.Message, nota = "Nessuna modifica è stata applicata: la transazione è stata annullata interamente." });
            }
        }

        // ------------------------------------------------------------------
        // Endpoint GET /apri: apre una pagina che non si trova nella cartella
        // dell'eseguibile, in due modi alternativi (mutuamente esclusivi):
        //
        //   /apri?percorso=C:\altra\cartella\pagina.htm
        //       Legge e serve un file ovunque sul disco locale o su una
        //       condivisione di rete (\\server\condivisa\...). A differenza di
        //       ServiFileStaticoAsync, qui il percorso NON viene limitato alla
        //       cartella dell'exe: è voluto, per poter aprire file altrove.
        //
        //   /apri?url=https://esempio.it/pagina
        //       Fa da proxy verso una pagina esterna su Internet: la richiede
        //       lato server e ne inoltra il contenuto (stesso Content-Type)
        //       così come arriva, permettendo di aprirla anche se il PC che
        //       usa il browser non ha accesso diretto a quell'indirizzo, o per
        //       incorporarla in una pagina servita da questo stesso server.
        // ------------------------------------------------------------------
        private static async Task HandleApriAsync(HttpListenerContext context)
        {
            var req = context.Request;
            var res = context.Response;

            string? percorso = req.QueryString["percorso"];
            string? url = req.QueryString["url"];

            if (!string.IsNullOrWhiteSpace(percorso))
            {
                await ServiFileAssolutoAsync(res, percorso);
                return;
            }

            if (!string.IsNullOrWhiteSpace(url))
            {
                await ServiPaginaEsternaAsync(res, url);
                return;
            }

            res.StatusCode = 400;
            await WriteJsonAsync(res, new
            {
                successo = false,
                errore = "Specifica il parametro 'percorso' (un file sul disco, anche fuori dalla cartella del programma) oppure 'url' (una pagina su Internet)."
            });
        }

        // Serve un file da un percorso assoluto qualsiasi, senza restringerlo alla
        // cartella dell'eseguibile: a differenza di ServiFileStaticoAsync, qui è
        // esplicitamente permesso uscire da quella cartella, perché l'endpoint /apri
        // serve proprio a raggiungere pagine altrove sul disco (o in rete).
        private static async Task ServiFileAssolutoAsync(HttpListenerResponse res, string percorso)
        {
            string percorsoCompleto;
            try
            {
                percorsoCompleto = Path.GetFullPath(percorso);
            }
            catch (Exception ex)
            {
                res.StatusCode = 400;
                await WriteJsonAsync(res, new { successo = false, errore = "Percorso non valido: " + ex.Message });
                return;
            }

            if (!File.Exists(percorsoCompleto))
            {
                res.StatusCode = 404;
                await WriteJsonAsync(res, new { successo = false, errore = $"File non trovato: {percorsoCompleto}" });
                return;
            }

            res.ContentType = ContentTypeDaEstensione(Path.GetExtension(percorsoCompleto));
            byte[] bytes = await File.ReadAllBytesAsync(percorsoCompleto);
            res.ContentLength64 = bytes.Length;
            await res.OutputStream.WriteAsync(bytes);
        }

        // Recupera una pagina esterna su Internet e ne inoltra il contenuto al browser
        // (proxy semplice, senza riscrivere link o script contenuti nella pagina).
        private static async Task ServiPaginaEsternaAsync(HttpListenerResponse res, string url)
        {
            if (!Uri.TryCreate(url, UriKind.Absolute, out Uri? uriEsterno) ||
                (uriEsterno.Scheme != Uri.UriSchemeHttp && uriEsterno.Scheme != Uri.UriSchemeHttps))
            {
                res.StatusCode = 400;
                await WriteJsonAsync(res, new { successo = false, errore = "URL non valido: deve iniziare con http:// o https://" });
                return;
            }

            try
            {
                using var rispostaEsterna = await _httpClient.GetAsync(uriEsterno);
                byte[] contenuto = await rispostaEsterna.Content.ReadAsByteArrayAsync();

                res.StatusCode = (int)rispostaEsterna.StatusCode;
                res.ContentType = rispostaEsterna.Content.Headers.ContentType?.ToString() ?? "application/octet-stream";
                res.ContentLength64 = contenuto.Length;
                await res.OutputStream.WriteAsync(contenuto);
            }
            catch (Exception ex)
            {
                // 502 Bad Gateway: il server ha provato a contattare la pagina esterna
                // ma non ha ottenuto una risposta valida (sito irraggiungibile, timeout, ecc.)
                res.StatusCode = 502;
                await WriteJsonAsync(res, new { successo = false, errore = "Impossibile raggiungere la pagina esterna: " + ex.Message });
            }
        }

        // Cerca nella cartella dell'exe tutti i file .accdb e .mdb, e indica quale
        // dei due è attualmente collegato al server (nessuno, se _db è nullo).
        private static async Task HandleListDatabasesAsync(HttpListenerContext context)
        {
            var res = context.Response;
            var dbAttuale = _db;

            try
            {
                var trovati = TrovaDatabaseNellaCartella().Select(percorso =>
                {
                    var info = new FileInfo(percorso);
                    return new
                    {
                        nomeFile = info.Name,
                        percorso = info.FullName,
                        dimensioneByte = info.Length,
                        ultimaModifica = info.LastWriteTime,
                        selezionato = dbAttuale != null && string.Equals(info.FullName, dbAttuale.DbPath, StringComparison.OrdinalIgnoreCase)
                    };
                }).ToList();

                await WriteJsonAsync(res, new { successo = true, cartella = _exeDir, database = trovati });
            }
            catch (Exception ex)
            {
                res.StatusCode = 500;
                await WriteJsonAsync(res, new { successo = false, errore = ex.Message });
            }
        }

        // Cambia il database attivo. Il percorso può essere assoluto oppure solo il nome
        // del file (in tal caso viene cercato nella cartella dell'exe). Se 'crea' non è
        // true e il file non esiste, l'operazione fallisce senza creare nulla.
        private static async Task HandleSelectDatabaseAsync(HttpListenerContext context)
        {
            var req = context.Request;
            var res = context.Response;

            var body = await ReadJsonAsync<DatabaseSelectRequest>(req);
            if (body == null || string.IsNullOrWhiteSpace(body.Percorso))
            {
                res.StatusCode = 400;
                await WriteJsonAsync(res, new { successo = false, errore = "Devi specificare il campo 'percorso'." });
                return;
            }

            string percorsoCompleto = Path.IsPathRooted(body.Percorso)
                ? body.Percorso
                : Path.Combine(_exeDir, body.Percorso);

            try
            {
                await EseguiOperazioneDbAsync(() => { ConnettiA(percorsoCompleto, creaSeMancante: body.Crea); return true; });
                await WriteJsonAsync(res, new { successo = true, percorso = percorsoCompleto });
            }
            catch (DbOccupatoException ex)
            {
                res.StatusCode = 503; // Service Unavailable: riprovare più tardi
                await WriteJsonAsync(res, new { successo = false, occupato = true, errore = ex.Message });
            }
            catch (Exception ex)
            {
                res.StatusCode = 400;
                await WriteJsonAsync(res, new { successo = false, errore = ex.Message });
            }
        }

        // Restituisce l'elenco delle tabelle del database attualmente connesso, con nome
        // e tipo di ciascun campo. Le chiavi JSON sono in camelCase (nome/colonne/tipo)
        // per coerenza con gli altri endpoint della pagina web.
        private static async Task HandleSchemaAsync(HttpListenerContext context)
        {
            var res = context.Response;

            var db = _db;
            if (db == null)
            {
                res.StatusCode = 400;
                await WriteJsonAsync(res, new { successo = false, errore = "Nessun database connesso." });
                return;
            }

            try
            {
                var risultato = await EseguiOperazioneDbAsync(() =>
                {
                    var tabelle = db.GetSchema().Select(t => new
                    {
                        nome = t.Nome,
                        colonne = t.Colonne.Select(c => new { nome = c.Nome, tipo = c.Tipo })
                    }).ToList();

                    // Query, macro e report vengono letti separatamente dalle tabelle: se uno di
                    // questi fallisce (es. ADOX non installato, o permesso negato su MSysObjects
                    // per macro/report), lo si segnala nel campo corrispondente senza far fallire
                    // l'intera risposta.
                    object query;
                    try
                    {
                        query = db.GetQueries().Select(q => new { nome = q.Nome, tipo = q.Tipo, sql = q.Sql });
                    }
                    catch (Exception ex)
                    {
                        query = new { errore = "Impossibile leggere le query: " + ex.Message };
                    }

                    object macro;
                    try
                    {
                        macro = db.GetMacro().Select(m => new { nome = m.Nome });
                    }
                    catch (Exception ex)
                    {
                        macro = new { errore = "Impossibile leggere le macro (permesso negato su MSysObjects): " + ex.Message };
                    }

                    object report;
                    try
                    {
                        report = db.GetReport().Select(r => new { nome = r.Nome });
                    }
                    catch (Exception ex)
                    {
                        report = new { errore = "Impossibile leggere i report (permesso negato su MSysObjects): " + ex.Message };
                    }

                    return new { tabelle, query, macro, report };
                });

                await WriteJsonAsync(res, new { successo = true, risultato.tabelle, risultato.query, risultato.macro, risultato.report });
            }
            catch (DbOccupatoException ex)
            {
                res.StatusCode = 503; // Service Unavailable: riprovare più tardi
                await WriteJsonAsync(res, new { successo = false, occupato = true, errore = ex.Message });
            }
            catch (Exception ex)
            {
                res.StatusCode = 500;
                await WriteJsonAsync(res, new { successo = false, errore = ex.Message });
            }
        }

        private static async Task<T?> ReadJsonAsync<T>(HttpListenerRequest req)
        {
            using var reader = new StreamReader(req.InputStream, req.ContentEncoding);
            string body = await reader.ReadToEndAsync();
            if (string.IsNullOrWhiteSpace(body)) return default;
            return JsonSerializer.Deserialize<T>(body, new JsonSerializerOptions { PropertyNameCaseInsensitive = true });
        }

        private static async Task WriteJsonAsync(HttpListenerResponse res, object data)
        {
            res.ContentType = "application/json; charset=utf-8";
            string json = JsonSerializer.Serialize(data);
            byte[] bytes = Encoding.UTF8.GetBytes(json);
            res.ContentLength64 = bytes.Length;
            await res.OutputStream.WriteAsync(bytes);
        }

        private static async Task WriteHtmlAsync(HttpListenerResponse res, string html)
        {
            res.ContentType = "text/html; charset=utf-8";
            byte[] bytes = Encoding.UTF8.GetBytes(html);
            res.ContentLength64 = bytes.Length;
            await res.OutputStream.WriteAsync(bytes);
        }

        // ------------------------------------------------------------------
        // Pagina home.htm: NON è incorporata nel programma. Viene letta da
        // disco (deve trovarsi nella stessa cartella dell'eseguibile) a ogni
        // richiesta, così l'exe non contiene alcuna copia della pagina web.
        // ------------------------------------------------------------------
        private static async Task ServiHomeHtmAsync(HttpListenerResponse res)
        {
            string percorso = Path.Combine(_exeDir, "home.htm");

            if (!File.Exists(percorso))
            {
                res.StatusCode = 404;
                res.ContentType = "text/plain; charset=utf-8";
                byte[] msg = Encoding.UTF8.GetBytes(
                    $"File 'home.htm' non trovato nella cartella dell'eseguibile ({_exeDir}). " +
                    "Copia il file home.htm accanto a MiniDbServer.exe.");
                res.ContentLength64 = msg.Length;
                await res.OutputStream.WriteAsync(msg);
                return;
            }

            string html = await File.ReadAllTextAsync(percorso, Encoding.UTF8);
            await WriteHtmlAsync(res, html);
        }

        // ------------------------------------------------------------------
        // Serve un qualsiasi file statico (html, css, js, csv, immagini, ecc.)
        // trovato nella cartella dell'eseguibile o in una sua sottocartella.
        // Restituisce true se il file è stato trovato e servito, false se non
        // esiste (così il chiamante può rispondere con il consueto 404 JSON).
        //
        // Sicurezza: il percorso richiesto viene risolto in un percorso assoluto
        // e si verifica che ricada davvero dentro _exeDir, per impedire di uscire
        // dalla cartella con sequenze tipo "..%2f..%2f" (path traversal).
        // ------------------------------------------------------------------
        private static async Task<bool> ServiFileStaticoAsync(HttpListenerResponse res, string urlPath)
        {
            // Decodifica eventuali caratteri percent-encoded (es. spazi come %20) e
            // normalizza le barre, poi toglie la barra iniziale per ottenere un
            // percorso relativo alla cartella dell'exe.
            string relativo = Uri.UnescapeDataString(urlPath).TrimStart('/', '\\');
            if (string.IsNullOrWhiteSpace(relativo)) return false;

            string radiceCompleta = Path.GetFullPath(_exeDir);
            string percorsoCompleto;
            try
            {
                percorsoCompleto = Path.GetFullPath(Path.Combine(radiceCompleta, relativo));
            }
            catch
            {
                return false; // percorso non valido
            }

            // Il file risolto deve restare dentro la cartella dell'exe (blocca "..").
            bool dentroLaCartella = percorsoCompleto.StartsWith(
                radiceCompleta + Path.DirectorySeparatorChar, StringComparison.OrdinalIgnoreCase)
                || string.Equals(percorsoCompleto, radiceCompleta, StringComparison.OrdinalIgnoreCase);

            if (!dentroLaCartella || !File.Exists(percorsoCompleto))
            {
                return false;
            }

            res.ContentType = ContentTypeDaEstensione(Path.GetExtension(percorsoCompleto));
            byte[] bytes = await File.ReadAllBytesAsync(percorsoCompleto);
            res.ContentLength64 = bytes.Length;
            await res.OutputStream.WriteAsync(bytes);
            return true;
        }

        // Content-Type minimo per i tipi di file più comuni; per tutto il resto
        // usa un tipo binario generico, che il browser gestisce comunque bene
        // per il download o la visualizzazione.
        private static string ContentTypeDaEstensione(string estensione)
        {
            return estensione.ToLowerInvariant() switch
            {
                ".htm" or ".html" => "text/html; charset=utf-8",
                ".css" => "text/css; charset=utf-8",
                ".js" => "application/javascript; charset=utf-8",
                ".json" => "application/json; charset=utf-8",
                ".csv" => "text/csv; charset=utf-8",
                ".txt" => "text/plain; charset=utf-8",
                ".png" => "image/png",
                ".jpg" or ".jpeg" => "image/jpeg",
                ".gif" => "image/gif",
                ".svg" => "image/svg+xml",
                ".ico" => "image/x-icon",
                _ => "application/octet-stream",
            };
        }
    }
}
