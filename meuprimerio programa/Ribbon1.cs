using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
namespace meuprimerio_programa
    {
    public partial class Ribbon1
        {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
            {

            }

        private void OlaMundoBtn_Click(object sender, RibbonControlEventArgs e)
            {


            Globals.ThisWorkbook.funções.Iniciar_Planilha_Nós();
            }

        private void Button1_Click(object sender, RibbonControlEventArgs e)
            {
            Globals.ThisWorkbook.funções.Iniciar_Ler_Pontos_dos_Nós();
            }

        private void Button2_Click(object sender, RibbonControlEventArgs e)
            {
            Globals.ThisWorkbook.funções.Iniciar_Planilha_Conectividade();
            }
        }
    }