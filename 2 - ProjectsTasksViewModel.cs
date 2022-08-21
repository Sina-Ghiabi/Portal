using Dapna.MSVPortal.Enums;
using Dapna.MSVPortal.ProjectsDocumentations;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Threading.Tasks;

namespace Dapna.MSVPortal.Web.ViewModels
{
    public class ProjectsTasksViewModel
    {
        public int? CreatorUserID { get; set; }

        //ProjectsTask Fields
        public int TaskID { get; set; }
        public int ProjectID { get; set; }
        public string ProjectName { get; set; }
        public string ProjectCode { get; set; }
        public string CompanyName { get; set; }
        public string DocumentTitle { get; set; }
        public string DocumentNumber { get; set; }
        public ProjectsTasksDiciplineTypes Dicipline { get; set; }
        public string ResponsiblePerson { get; set; }
        public ProjectsTasksDocumentTypes? DocumentType { get; set; }
        public string Description { get; set; }

        //ProjectsTaskSchedule Fields
        public int? WeightFactor { get; set; }
        public int? Progress { get; set; }
        public string BaseLineStart { get; set; }
        public string BaseLineFinished { get; set; }
        public string PlanStart { get; set; }
        public string PlanFinished { get; set; }
        public string ActualStart { get; set; }
        public string ActualFinished { get; set; }
        public int? OriginalDuration { get; set; }
        public ProjectsTasksSourceOfItemTypes? SourceOfItem { get; set; }
        public int? ManPower { get; set; }
        public bool Critical { get; set; }

        //Fields Below Will Be Filled From Last Revision 
        public string LastTransmitalNumber { get; set; }
        public string LastTransmitalDate { get; set; }
        public string LastRevisionNumber { get; set; }
        public ProjectsTasksStatusTypes? LastStatus { get; set; }
        public ProjectsTasksActionTypes? LastAction { get; set; }

        //Dashboard 
        
        public delegate void AverageAFCDocuments(List<ProjectsTasks> AverageAFCDocuments , int DiciplineType, int? ProjectID, DateTime? StartDate , DateTime? EndDate);

        public delegate void DiciplineProduceProcess(List<ProjectsTasksRevisions> List , int? ImportYear = null , int? ImportStartMonth = null , int? ImportEndMonth = null);
        public SelectiveYear DefaultYear { get; set; }
        public SelectiveMonth DefaultMonth { get; set; }
        public int AverageProjectsDocumentations { get; set; }
        public string StartDate { get; set; }
        public string EndDate { get; set; }

        //Anonymous Functions
        public int? CountTransmitalNumber { get; set; }

        public delegate string DateConverter(string Date);

        public delegate int ConvertAction(string Action);

        public delegate string EnglishDigitToPersian(string Date);

        public delegate string PersianDigitToEnglish(string Date);

        public delegate string ConvertToGregorian(string Date);

        public delegate int EnumDataChecker(int Date);

        public delegate List<ProjectsTasksViewModel> ConnectToDatabase(string Query);

        //Count Number Of Transmitals
        public int? CountTransmitals { get; set; }
    }
}
