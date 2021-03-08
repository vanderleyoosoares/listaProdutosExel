using System;
using System.Collections.Generic;
using System.Linq;
using ClosedXML.Excel;

namespace roupas
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Escaneie os códigos de barra:");
            int x ;
            string v;
            bool n = true;
            List<int> myLS = new List<int>();
            while( n == true )
            {
                v = Console.ReadLine();
                if(v == "")
                    {
                        n = false;
                        break;
                    }
                x = Convert.ToInt32(v);
                myLS.Add(x);
            }

            myLS.Sort();

            //contagem dos itens e criação de lista com contagem:
            List<int> quantidade = new List<int>();
            List<int> listOk = new List<int>(myLS.Distinct());
            int qtd = 0;
            for(int l = 0; l  < listOk.Count; l++)
            {
                qtd = 0;
                for(int c = 0; c < myLS.Count; c++)
                {
                    if(listOk[l] == myLS[c])
                    {
                        qtd++;
                    }
                }
                quantidade.Add(qtd);
            }

//-----------------------------------------------------------------------------------------------------------------

            var wb = new XLWorkbook(@"C:\projetoroupas\controle.xlsx");
            var planilha = wb.Worksheet(1);
            
            for(int tl = 0; tl < listOk.Count; tl++)
            {
                string cUm = "A"+(tl+2);
                string cDois = "B"+(tl+2);

                planilha.Cell(cUm).Value = listOk[tl];
                planilha.Cell(cDois).Value = quantidade[tl];

            }
            wb.SaveAs(@"C:\projetoroupas\controle.xlsx");
//-----------------------------------------------------------------------------------------------------------------
        }
       
    }
}