using Abp.Application.Services.Dto;
using Abp.AutoMapper;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Text;

namespace Dapna.MSVPortal.Financial.Dto
{
    [AutoMap(typeof(RemainCredit))]
    public class RemainCreditDto : FullAuditedEntityDto<int>
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
        public long? RemainDebtAmount { get; set; }
        public long? RemainCreditAmount { get; set; }
    }
}
