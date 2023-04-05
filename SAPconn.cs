using SAP.Middleware.Connector;
using System;
using System.Collections.Generic;
using System.Data;

namespace KSAPCon {

    /// <summary>
    /// SAP Connector 3.1
    /// </summary>
    public class SAPconn {

        #region Constructors

        /// <summary>
        /// SAP Connector 3.1 global User
        /// </summary>
        public SAPconn() {
            SapCon3("CADUSER", "trz100");
        }

        /// <summary>
        /// SAP Connector 3.1
        /// </summary>
        /// <param name="user">K-Nummer des Besnutzers</param>
        /// <param name="password">SAP Passwort</param>
        public SAPconn(string user, string password) {
            SapCon3(user, password);
        }

        #endregion

        #region Properties

        public bool ConnOK { get; set; }
        private RfcDestination _con { get; set; }

        #endregion

        #region Methods

        /// <summary>
        /// Konfiguration´aus lesen
        /// </summary>
        /// <param name="beleg"></param>
        /// <param name="belegnummer"></param>
        /// <returns></returns>
        public Dictionary<string, confItem> config(string beleg, string belegnummer) {
            var dic = new Dictionary<string, confItem>();
            var repo = _con.Repository;

            var fnc = repo.CreateFunction("Z_CLAF_CLASS_OF_OBJECTS");
            fnc.SetValue("VBELN", beleg);
            fnc.SetValue("POSNR", belegnummer);
            fnc.SetValue("INITIAL_CHARACT", " ");
            fnc.SetValue("NO_VALUE_DESCRIPT", "X");
            fnc.SetValue("DESC_TO_AUSP2", "X");
            //if (cuobj != "")
            //    fnc.SetValue("CUOBJ", cuobj);

            //call function module
            fnc.Invoke(_con);
            //work with response
            var tab = fnc.GetTable("T_OBJECTDATA");
            for (var i = 0; i < tab.Count; i++) {
                var row = tab[i];

                var conf = new confItem();

                conf.name = row.GetString("ATNAM");
                conf.val1 = row.GetString("AUSP1");
                conf.val2 = row.GetString("AUSP2");
                conf.dsr = row.GetString("SMBEZ");
                conf.id = row.GetString("ATIMB");

                if (!dic.ContainsKey(conf.name)) {
                    dic.Add(conf.name, conf);
                }
            }
            return dic;
        }

        /// <summary>
        /// aktuelle PlanVersion ermitteln
        /// </summary>
        /// <param name="plannummer"></param>
        /// <returns></returns>
        public Dictionary<string, string> PlanVersion(string plannummer) {
            var cut = plannummer.Length - 2;
            var plan = plannummer.Substring(0, cut) + "00";

            var repo = _con.Repository;

            //var cmd = string.Format("RH_LOCAL_INBOX_GET, Us)
            var fnc = repo.CreateFunction("Z_AV_GET_OBJECT_DOCUMENTFILES");
            fnc.SetValue("DOCUMENTNUMBER", plan);
            fnc.SetValue("CURRENT_VERSION", "X");

            fnc.Invoke(_con);

            var table = fnc.GetTable("DOCUMENTFILES");

            var row = table[table.Count - 1];

            var nummer = row.GetValue("DOCUMENTNUMBER").ToString();
            var version = row.GetValue("DOCUMENTVERSION").ToString();
            var part = row.GetValue("DOCUMENTPART").ToString();
            var typ = row.GetValue("DOCUMENTTYPE").ToString();
            var file = row.GetValue("DOCFILE").ToString();
            var app = row.GetValue("WSAPPLICATION").ToString();

            ////https://autovue.krones-deu.krones-group.com:5001/?&DIRKEY=PZ2@PL150999A0000004@00@000@WS=MED

            //var url1 = string.Format(@"http://autovue.krones-deu.krones-group.com:5000/?&DIRKEY={0}@{1}@{2}@{3}@WS={4}",
            //   typ, nummer, version, part, app);

            //Process.Start("iexplore.exe", url1);
            var planversion = new Dictionary<string, string> {
                { "typ", typ },
                { "nummer", nummer },
                { "version", version },
                { "part", part },
                { "app", app }
            };

            return planversion;
        }

        public DataTable VBAP(string vbeln, string pos) {
            try {
                var repo = _con.Repository;
                var fnc = repo.CreateFunction("Z_RFC_VBAP_SUCHEN");

                fnc.SetValue("BO_NEU", "X");
                fnc.SetValue("VBELN", vbeln);
                fnc.SetValue("POSNR", pos);
                //call function module
                fnc.Invoke(_con);
                //work with response
                var vbap_tab = fnc.GetTable("VBAP_NEU");

                var dt = ConvertToDoNetTable(vbap_tab);

                //string[] sZeile = new string[10];
                //int anz = vbap_tab.Count;
                //if (anz > 0) {
                //    for (int i = 0; i < anz; i++) {
                //        var metadata = vbap_tab.GetElementMetadata(i);
                //        var vbap_zeile = vbap_tab[i];

                //        sZeile[0] = vbap_zeile.GetString("VBELN");
                //        sZeile[1] = vbap_zeile.GetString("POSNR");
                //        sZeile[2] = vbap_zeile.GetString("ARKTX");
                //        sZeile[3] = vbap_zeile.GetString("ZZ_EQUI");
                //        if (sZeile[3] == "")
                //            sZeile[3] = vbap_zeile.GetString("ZZ_EQUIART");
                //        sZeile[4] = vbap_zeile.GetString("SERNR");
                //        sZeile[5] = vbap_zeile.GetString("ZZ_ANLAGE");
                //        sZeile[6] = vbap_zeile.GetString("OBJNR");
                //        sZeile[7] = vbap_zeile.GetString("MATNR");
                //        sZeile[8] = vbap_zeile.GetString("ZZ_MATMOD");
                //        //vbeln_aus.Tabelle.Items.Add(new ListViewItem(sZeile));
                //    }
                //}
                return dt;
            } catch {
                return null;
            }
        }

        public vbel Vetriebsbeleg(string vbeln, string posnr) {
            var repo = _con.Repository;
            var fnc = repo.CreateFunction("Z_RFC_KOMNR_INFO");

            fnc.SetValue("VBELN", vbeln);
            fnc.SetValue("POSNR", posnr);

            fnc.Invoke(_con);
            //work with response
            var vb = new vbel();

            vb.komm = fnc.GetString("OKOMNR");
            vb.vbeln = fnc.GetString("OVBELN");
            vb.posnr = fnc.GetString("OPOSNR");
            vb.deb_no = fnc.GetString("AUFTRAGGEBER");
            vb.deb_txt = fnc.GetString("AUFTRAGGEBERTXT");
            vb.rec_no = fnc.GetString("WARENEMP");
            vb.rec_txt = fnc.GetString("WARENEMPTXT");
            vb.laneNo = fnc.GetString("ZZ_ANLAGE");
            vb.netplan = fnc.GetString("OAUFNR");
            vb.model = fnc.GetString("ZZ_MATMOD");
            vb.modellarge = fnc.GetString("ARKTX");
            vb.typ = fnc.GetString("ZZ_EQUIART");
            vb.typ1 = fnc.GetString("EARTX");
            vb.info = fnc.GetString("BO_VBAP_INFO");

            return vb;
        }

        /// <summary>
        /// aktuelle Wokflow des Users
        /// </summary>
        /// <param name="user_ID">KNummer des Users</param>
        /// <returns></returns>
        public DataTable WorkFlow(string user_ID) {
            //    CLIENT Type MANDT CLNT 3    0    Mandant
            //WI_ID Type SWW_WIID NUMC    12    0    Workitem-Kennung
            //WI_TYPE Type SWW_WITYPE CHAR    1    0    Workitem-Typ
            //WI_CREATOR Type SWW_OBJID CHAR    90    0    Erzeuger eines Workitems
            //WI_LANG Type SWW_LANG LANG    1    0    Sprache für Texte eines Workitem
            //WI_TEXT Type WITEXT CHAR    120    0    Workitem-Text
            //WI_NITEXT Type SWW_NITEXT CHAR    120    0    Kurztext für die Endebenachrichtigung zum Workitem
            //WI_DITEXT Type SWW_DITEXT CHAR 120    0    Kurztext für die Terminüberschreitung zum Workitem
            //WI_RHTEXT Type SWW_RHTEXT CHAR 40    0    Kurztext aus Aufgabe
            //WI_CD Type SWW_CD DATS    8    0    Erzeugungsdatum eines Workitem
            //WI_CT Type SWW_CT TIMS    6    0    Erzeugungszeit eines Workitem
            //WI_AAGENT Type SWW_AAGENT CHAR    12    0    Tatsächlicher Bearbeiter eines Workitems
            //WI_CHCKWI Type SWW_CHCKWI NUMC 12    0    Übergeordnetes Workitem
            //WI_CBFB Type SWW_CBFB CHAR 30    0    Callback Funktionsbaustein für Rückmeldung eines Workitems
            //CB_DONE Type SWW_CBDONE CHAR 1    0    Statusübergang für ausgeführtes Workitem durchgeführt
            //WI_RH_TASK Type SWW_TASK CHAR    14    0    Aufgabenkennung
            //WI_PRIO Type SWW_PRIO NUMC    1    0    Priorität eines Workitem
            //WI_CONFIRM Type SWW_WICONF CHAR    1    0    Kennzeichen, daß Bearbeitungsende bestätigt werden muß
            //WI_COMP_EV Type SWW_COMPEV CHAR    1    0    Kennzeichen, daß Beendigung durch Ereignis möglich
            //WI_REJECT Type SWW_REJECT CHAR    1    0    Kennzeichen, daß das WI vom Bearbeiter zurückweisbar ist
            //WI_FORW_BY Type SWW_FORWBY CHAR    12    0    (letzter) Weiterleitender eines Workitems
            //WI_DISABLE Type SWW_DISABL CHAR 1    0    Kennzeichen, daß das Workitem z.Zt.gesperrt ist
            //STATUSTEXT Type SWW_STATXT CHAR 20    0    Workflow: Workitem-Status
            //    WI_STAT Type SWW_WISTAT CHAR 12    0    Bearbeitungsstatus eines Workitem
            //WI_LED Type SWW_LED DATS 8    0    Datum der Frist des Workitem
            //WI_LET Type SWW_LET TIMS 6    0    Zeit zur Frist des Workitems
            //TYPETEXT Type SWW_TYPTXT CHAR 20    0    Typbezeichnung eines Workitem
            //TCLASS Type HR_TASK_CL CHAR 12    0    Klassifikation von Aufgaben
            //NOTE_EXIST Type SWW_NOTEEX CHAR 1    0    Kennzeichen, daß zu dem WI mindestens eine Notiz existiert
            //WI_DH_STAT Type SWW_WIDHST NUMC    4    0    Deadline Status eines Workitems
            //DHSTATEXT Type SWW_STATXT CHAR    20    0    Workflow: Workitem-Status

            var repo = _con.Repository;
            var lang = _con.Language.Substring(0, 1);
            //var cmd = string.Format("RH_LOCAL_INBOX_GET, Us)
            var funktion = repo.CreateFunction("RH_LOCAL_INBOX_GET");
            funktion.SetValue("USER_ID", user_ID);
            funktion.SetValue("USER_LANGU", lang);
            funktion.Invoke(_con);

            var table = funktion.GetTable("WI_LIST");

            var dt = ConvertToDoNetTable(table);
            return dt;
        }

        /// <summary>
        /// Konvertiert RFCTable -> Datatabel
        /// </summary>
        /// <param name="rfcTable">IRfcTable Tabelle</param>
        /// <returns>Daten Tabellen</returns>
        private DataTable ConvertToDoNetTable(IRfcTable rfcTable) {
            var dt = new DataTable();

            // Colums der Tab
            for (var item = 0; item < rfcTable.ElementCount; item++) {
                var metadata = rfcTable.GetElementMetadata(item);
                dt.Columns.Add(metadata.Name);
            }

            foreach (var row in rfcTable) {
                var datarow = dt.NewRow();

                for (var item = 0; item < rfcTable.ElementCount; item++) {
                    var metadata = rfcTable.GetElementMetadata(item);
                    if (metadata.DataType == RfcDataType.BCD && metadata.Name == "ABC") {
                        datarow[item] = row.GetInt(metadata.Name);
                    } else {
                        datarow[item] = row.GetString(metadata.Name);
                    }
                }
                dt.Rows.Add(datarow);
            }

            return dt;
        }

        /// <summary>
        /// SAP Connection erstellen
        /// </summary>
        /// <param name="user">K Nummer</param>
        /// <param name="password">SAP Passwort</param>
        private void SapCon3(string user, string password) {
            try {
                var sap_conf = new RfcConfigParameters {
                    //sap_conf.Add(RfcConfigParameters.Language, SAPdest.Language);
                    { RfcConfigParameters.Language, "DE" },
                    { RfcConfigParameters.UseSAPGui, "1" },
                    { RfcConfigParameters.AbapDebug, "1" },
                    //sap_conf.Add(RfcConfigParameters.PoolSize, "5");
                    //sap_conf.Add(RfcConfigParameters.MaxPoolSize, "10");
                    { RfcConfigParameters.ConnectionIdleTimeout, "600" },
                    { RfcConfigParameters.MessageServerHost, "erp-p-ntr002" },
                    { RfcConfigParameters.SystemID, "KP1" },
                    { RfcConfigParameters.Name, "KP1" },
                    { RfcConfigParameters.LogonGroup, "PUBLIC" },
                    { RfcConfigParameters.Client, "100" },
                    { RfcConfigParameters.User, user },
                    { RfcConfigParameters.Password, password }
                };
                _con = RfcDestinationManager.GetDestination(sap_conf);                 ////connect.AbapDebug = "1";

                ConnOK = true;
            } catch {
                ConnOK = false;
            }
        }

        #endregion
    }
}