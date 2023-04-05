using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace KSAPCon {

    /// <summary>
    /// Vertriebsbeleg
    /// </summary>
    public class vbel {

        #region Properties

        /// <summary>
        /// Kundennummer
        /// </summary>
        public string deb_no { get; set; }

        /// <summary>
        /// Kundenname
        /// </summary>
        public string deb_txt { get; set; }

        /// <summary>
        /// Information
        /// </summary>
        public string info { get; set; }

        /// <summary>
        /// Kommisionnummer
        /// </summary>
        public string komm { get; set; }

        /// <summary>
        /// Anlagennummer
        /// </summary>
        public string laneNo { get; set; }

        /// <summary>
        /// Maschinenmodel
        /// </summary>
        public string model { get; set; }

        /// <summary>
        /// Maschinenmodel lang
        /// </summary>
        public string modellarge { get; set; }

        /// <summary>
        /// Netzplannummer
        /// </summary>
        public string netplan { get; set; }

        /// <summary>
        /// Vertreibesbelegposition
        /// </summary>
        public string posnr { get; set; }

        /// <summary>
        /// Wahrenempfänger nummer
        /// </summary>
        public string rec_no { get; set; }

        /// <summary>
        /// Wahrenempfängername
        /// </summary>
        public string rec_txt { get; set; }

        /// <summary>
        /// Maschinentyp
        /// </summary>
        public string typ { get; set; }

        public string typ1 { get; set; }

        /// <summary>
        /// Vertriebsbelegnummer
        /// </summary>
        public string vbeln { get; set; }

        #endregion

        /// <summary>
        /// Maschinentyp
        /// </summary>
    }
}