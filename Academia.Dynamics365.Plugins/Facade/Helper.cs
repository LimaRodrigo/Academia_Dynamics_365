using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace Academia.Dynamics365.Backend.Facade
{
    internal static class Helper
    {
        public static bool ValidarCpf(string cpf)
        {
            if (string.IsNullOrEmpty(cpf)) return false;

            int[] d = new int[14];
            int[] v = new int[2];
            int j, i, soma;
            string CpfSoNumero = Regex.Replace(cpf, "[^0-9]", string.Empty);

            if (new string(CpfSoNumero[0], CpfSoNumero.Length) == CpfSoNumero) return false;

            if (CpfSoNumero.Length == 11)
            {
                for (i = 0; i <= 10; i++) d[i] = Convert.ToInt32(CpfSoNumero.Substring(i, 1));
                for (i = 0; i <= 1; i++)
                {
                    soma = 0;
                    for (j = 0; j <= 8 + i; j++) soma += d[j] * (10 + i - j);

                    v[i] = (soma * 10) % 11;
                    if (v[i] == 10) v[i] = 0;
                }
                return (v[0] == d[9] & v[1] == d[10]);
            }
            else return false;

        }
    }
}
