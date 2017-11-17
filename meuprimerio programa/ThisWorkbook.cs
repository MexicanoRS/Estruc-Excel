using System;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using Microsoft.VisualStudio.Tools.Applications.Runtime;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;



namespace meuprimerio_programa
	{

	public partial class ThisWorkbook
		{

		public Funções funções;
		public Excel.Worksheet Nós;
		public Excel.Worksheet Conectividade;
		private void ThisWorkbook_Startup( object sender, System.EventArgs e )
			{
			Nós = ( Excel.Worksheet )Globals.ThisWorkbook.Worksheets.Add();
			Nós.Visible = Excel.XlSheetVisibility.xlSheetVeryHidden;
			Nós.Name = "Nós";
			Conectividade = ( Excel.Worksheet )Globals.ThisWorkbook.Worksheets.Add();
			Conectividade.Visible = Excel.XlSheetVisibility.xlSheetVeryHidden;
			Conectividade.Name = "Conectividade";



			funções = new Funções( Nós, Conectividade );
			}

		private void ThisWorkbook_Shutdown( object sender, System.EventArgs e )
			{
			}

		#region Código gerado pelo Designer VSTO

		/// <summary>
		/// Método necessário para suporte ao Designer - não modifique 
		/// o conteúdo deste método com o editor de código.
		/// </summary>
		private void InternalStartup( )
			{
			this.Startup += new System.EventHandler( ThisWorkbook_Startup );
			this.Shutdown += new System.EventHandler( ThisWorkbook_Shutdown );
			}

		#endregion

		}
	}
