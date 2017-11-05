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

            Excel.Worksheet Nós = (Excel.Worksheet)Globals.ThisWorkbook.Worksheets.Add();
            Nós.Name = "Nós";
            Nós.get_Range("A1").Value = "Preencha a tabela abaixo de conectividade:";
            Nós.Cells[ 2, 1 ].Value = "Nós";
            Nós.Cells[ 2, 2 ].Value = "X";
            Nós.Cells[ 2, 3 ].Value = "Y";
            string alfabeto = "abcdefghijklmnopqrstuvwxyz";
            int pos = 0;
            for ( int i = 3; i < 203; i++ )
                {
                if ( pos <= 25 )
                    {
                    Nós.Cells[ i, 1 ].Value = String.Concat(alfabeto[ pos ]);
                    }
                else
                    {
                    Nós.Cells[ i, 1 ].Value = String.Concat(alfabeto[ (pos / 26 )-1], alfabeto[ pos - 26 * ( pos / 26 ) ]);
                    }
                pos++;

                }
            Excel.Range Tabela_de_Nós = Nós.Range[Nós.Cells[ 3, 1 ], Nós.Cells[ 99, 3 ]];
            Tabela_de_Nós.Style.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            Nós.Cells[ 1, 1 ].HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
            }
        }
    }
