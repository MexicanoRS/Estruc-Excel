using System;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using Microsoft.VisualStudio.Tools.Applications.Runtime;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;

namespace meuprimerio_programa
    {
    public partial class Planilha2
        {
        private void Planilha2_Startup(object sender, System.EventArgs e)
            {
            }

        private void Planilha2_Shutdown(object sender, System.EventArgs e)
            {
            }

        #region Código gerado pelo Designer VSTO

        /// <summary>
        /// Método necessário para suporte ao Designer - não modifique 
        /// o conteúdo deste método com o editor de código.
        /// </summary>
        private void InternalStartup()
            {
            this.Startup += new System.EventHandler(Planilha2_Startup);
            this.Shutdown += new System.EventHandler(Planilha2_Shutdown);
            }

        #endregion

        }
    }
