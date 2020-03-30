using System;
using ExcelDna.Integration;
using System.Windows.Forms;

namespace ExcelInterface
{
    public static class ExcelInterface
    {
        public static void Execute()
        {
            try
            {
                // Creer un objet Excel application
                dynamic Excel = ExcelDnaUtil.Application;

                // Lecture de valeurs a partir de range XL
                double S = (double)Excel.Range["_S"].Value2;
                double K = (double)Excel.Range["_K"].Value2;
                double r = (double)Excel.Range["_r"].Value2;
                double v = (double)Excel.Range["_v"].Value2;
                double T = (double)Excel.Range["_T"].Value2;
                double t = (double)Excel.Range["_t"].Value2;

                // <Inserer code Pricer> 


                // Ecriture de valeurs dans un range XL
                double C = 1;
                Excel.Range["_C"] = C;

                double P = 0;
                Excel.Range["_P"] = P;

                double delta = 0;
                Excel.Range["_delta"] = delta;

                double theta = 0;
                Excel.Range["_theta"] = theta;

                double gamma = 0;
                Excel.Range["_gamma"] = gamma;

                double vega = 0;
                Excel.Range["_vega"] = vega;

                double rho = 0;
                Excel.Range["_rho"] = rho;
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message.ToString());
            }
        }
    }
}
