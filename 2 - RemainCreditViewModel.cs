using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Threading.Tasks;

namespace Dapna.MSVPortal.Web.ViewModels
{
    public class RemainCreditViewModel
    {
        [Required]
        public int ProjectID { get; set; }
        [Required]
        public string ProjectCode { get; set; }
        [Required]
        public string ProjectName { get; set; }
        [Required]
        public string FourthLevelCode { get; set; }
        [Required]
        public string PayTo { get; set; }
        public string RemainDebtAmount { get;set;}
        public string RemainCreditAmount { get; set; }
        public string UploadDate { get; set; }
        public string TotalDebtAmount { get; set; }
        public string TotalCreditAmount { get; set; }

    }
}
