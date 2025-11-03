using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EtatsComptables
{
    public class Grand_Livre
    {
		public int ID_ETATCMP { get; set; }
        public string Compte { get; set; }
        public string NomCompte { get; set; }
        public string Date { get; set; }
        public string Type_Numéro_de_pièce { get; set; }
        public string Libellé { get; set; }
        public string Libellé_ligne { get; set; }
        public string Date_Echeance { get; set; }
        public string Tiers { get; set; }
        public string Site { get; set; }
        public string Journal { get; set; }
        public string Lettre { get; set; }
        public string Etat { get; set; }
        public string Débit { get; set; }
        public string Crédit { get; set; }
        public string Solde_débiteur { get; set; }
        public string Solde_créditeur { get; set; }
    }
}
