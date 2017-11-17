using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Tools.Ribbon;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;


namespace meuprimerio_programa
	{
	public class Funções
		{

		public struct Tipo_Ponto
			{
			public Tipo_Ponto( string nome, double x, double y )
				{
				X = x;
				Y = y;
				Nome = nome;
				}

			public double X { get; set; }

			public double Y { get; set; }

			public string Nome { get; set; }
			}
		public struct Tipo_Barra
			{
			public Tipo_Barra( string nome, Tipo_Ponto nó_1, Tipo_Ponto nó_2 )
				{
				Nó = new Tipo_Ponto[ 2 ] { nó_1, nó_2 };
				Nome = nome;
				}

			public Tipo_Ponto[] Nó { get; set; }

			public string Nome { get; set; }
			}
		List<Tipo_Ponto> Pontos;
		internal Excel.Worksheet Nós;
		internal Excel.Worksheet Conectividade;
		public Funções( Excel.Worksheet nos, Excel.Worksheet conectividade )
			{
			Nós = nos;
			Conectividade = conectividade;
			}
		public void Iniciar_Planilha_Nós( )
			{
			Nós.get_Range( "A1" ).Value = "Preencha a tabela abaixo de conectividade:";
			Nós.Cells[ 2, 1 ].Value = "Nós";
			Nós.Cells[ 2, 2 ].Value = "X";
			Nós.Cells[ 2, 3 ].Value = "Y";
			string alfabeto = "abcdefghijklmnopqrstuvwxyz";
			int pos = 0;
			for ( int i = 3 ; i < 203 ; i++ )
				{
				if ( pos <= 25 )
					{
					Nós.Cells[ i, 1 ].Value = String.Concat( alfabeto[ pos ] );
					}
				else
					{
					Nós.Cells[ i, 1 ].Value = String.Concat( alfabeto[ ( pos / 26 ) - 1 ], alfabeto[ pos - 26 * ( pos / 26 ) ] );
					}
				pos++;

				}
			Excel.Range Tabela_de_Nós = Nós.Range[ Nós.Cells[ 3, 1 ], Nós.Cells[ 99, 3 ] ];
			Tabela_de_Nós.Style.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
			Nós.Cells[ 1, 1 ].HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
			Nós.Visible = Excel.XlSheetVisibility.xlSheetVisible;
#if DEBUG
            Nós.Cells[ 3, 2 ].Value = 0;
            Nós.Cells[ 3, 3 ].Value = 0;
            Nós.Cells[ 4, 2 ].Value = 5;
            Nós.Cells[ 4, 3 ].Value = 0;
            Nós.Cells[ 5, 2 ].Value = 5;
            Nós.Cells[ 5, 3 ].Value = 5;
            Nós.Cells[ 6, 2 ].Value = 8;
            Nós.Cells[ 6, 3 ].Value = 5;
#endif
			Nós.Activate();
			}


		public void Iniciar_Ler_Pontos_dos_Nós( )
			{
			Pontos = new List<Tipo_Ponto>();
			bool Parar_Leitura = false;
			int cont = 3;
			int sucessivos_nulos = 0;
			int Número_de_Tentativas_até_parar = 10;
			while ( Parar_Leitura == false )
				{
				if ( Nós.Cells[ cont, 2 ].Value != null || Nós.Cells[ cont, 3 ].Value != null )
					{
					sucessivos_nulos = 0;

					Pontos.Add( new Tipo_Ponto(
												( string )Nós.Cells[ cont, 1 ].Value,
												( double )Nós.Cells[ cont, 2 ].Value,
												( double )Nós.Cells[ cont, 3 ].Value ) );
					}
				else
					{
					sucessivos_nulos++;
					}
				cont++;
				if ( sucessivos_nulos == Número_de_Tentativas_até_parar ) Parar_Leitura = true;
				}
			System.Windows.Forms.MessageBox.Show( Pontos.ToString() );
			}

		public void Iniciar_Planilha_Conectividade( )
			{
			Conectividade.get_Range( "A1" ).Value = "Preencha a tabela abaixo de conectividade:";
			Conectividade.Cells[ 2, 1 ].Value = "Barras";
			Conectividade.Cells[ 2, 2 ].Value = "Nó Inicial";
			Conectividade.Cells[ 2, 3 ].Value = "Nó Final";
			List<string> Lista_de_nós = new List<string>();
			foreach ( Tipo_Ponto Ponto_em_questão in Pontos ) Lista_de_nós.Add( Ponto_em_questão.Nome );
			string ListaPontos = string.Join( ";", Lista_de_nós.ToArray() );
			string alfabeto = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
			int pos = 0;
			for ( int i = 3 ; i < 203 ; i++ )
				{
				if ( pos <= 25 )
					{
					Conectividade.Cells[ i, 1 ].Value = String.Concat( alfabeto[ pos ] );
					}
				else
					{
					Conectividade.Cells[ i, 1 ].Value = String.Concat( alfabeto[ ( pos / 26 ) - 1 ], alfabeto[ pos - 26 * ( pos / 26 ) ] );
					}
				pos++;
				}
			Excel.Range cell = Conectividade.Range[ Conectividade.Cells[ 3, 2 ], Conectividade.Cells[ 203, 3 ] ];
			cell.Validation.Delete();
			cell.Validation.Add(
				XlDVType.xlValidateList,
				XlDVAlertStyle.xlValidAlertInformation,
				XlFormatConditionOperator.xlBetween,
				ListaPontos,
				Type.Missing );
			cell.Validation.IgnoreBlank = true;
			cell.Validation.InCellDropdown = true;
			Excel.Range Tabela_de_Conectividade = Conectividade.Range[
																		Conectividade.Cells[ 3, 1 ],
																		Conectividade.Cells[ 99, 3 ] ];
			Tabela_de_Conectividade.Style.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
			Conectividade.Cells[ 1, 1 ].HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
			Conectividade.Visible = Excel.XlSheetVisibility.xlSheetVisible;
			Conectividade.Activate();
			}
		}
	}
