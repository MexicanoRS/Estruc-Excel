using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Tools.Ribbon;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;

namespace meuprimerio_programa
    {
    public class Funções
        {

        public struct Tipo_Ponto
            {
            public Tipo_Ponto(string nome, double x, double y)
                {
                X = x;
                Y = y;
                Nome = nome;
                }

            public double X { get; }

            public double Y { get; }

            public string Nome { get; }
            }

        List<Tipo_Ponto> Pontos = new List<Tipo_Ponto>();
        internal Excel.Worksheet Nós;

        public Funções(Excel.Worksheet nos)
            {
            Nós = nos;
            }
        public void Iniciar_Planilha_Nós()
            {
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
                    Nós.Cells[ i, 1 ].Value = String.Concat(alfabeto[ ( pos / 26 ) - 1 ], alfabeto[ pos - 26 * ( pos / 26 ) ]);
                    }
                pos++;

                }
            Excel.Range Tabela_de_Nós = Nós.Range[ Nós.Cells[ 3, 1 ], Nós.Cells[ 99, 3 ] ];
            Tabela_de_Nós.Style.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            Nós.Cells[ 1, 1 ].HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
            Nós.Visible = Excel.XlSheetVisibility.xlSheetVisible;
            Nós.Activate();
            }
        public void Iniciar_Ler_Pontos_dos_Nós()
            {
            bool Parar_Leitura = false;
            int cont = 3;
            int sucessivos_nulos = 0;
            int Número_de_Tentativas_até_parar = 10;
            while ( Parar_Leitura == false )
                {
                if ( Nós.Cells[ cont, 2 ].value2 != null|| Nós.Cells[ cont, 3 ].Value2 != null )
                    {
                    sucessivos_nulos = 0;
                    Pontos.Add(new Tipo_Ponto((string)Nós.Cells[ cont, 1 ].Value, (double)Nós.Cells[ cont, 2 ].Value, (double)Nós.Cells[ cont, 3 ].Value));

                    }
                else
                    {
                    sucessivos_nulos++;
                    }
                cont++;
                if ( sucessivos_nulos == Número_de_Tentativas_até_parar ) Parar_Leitura = true;
            
                }
            System.Windows.Forms.MessageBox.Show(Pontos.ToString());
            }




        }
    }
