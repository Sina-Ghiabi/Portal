using Dapna.MSVPortal.Projects.Dto;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Threading.Tasks;

namespace Dapna.MSVPortal.Web.ViewModels
{
    public class ChosenColumnsViewModel 
    {
        public int UserID { get; set; }
        public string Section { get; set; }
        public int TableID { get; set; }
        public bool Column1 { get; set; }
        public bool Column2 { get; set; }
        public bool Column3 { get; set; }
        public bool Column4 { get; set; }
        public bool Column5 { get; set; }
        public bool Column6 { get; set; }
        public bool Column7 { get; set; }
        public bool Column8 { get; set; }
        public bool Column9 { get; set; }
        public bool Column10 { get; set; }
        public bool Column11 { get; set; }
        public bool Column12 { get; set; }
        public bool Column13 { get; set; }
        public bool Column14 { get; set; }
        public bool Column15 { get; set; }
        public bool Column16 { get; set; }
        public bool Column17 { get; set; }
        public bool Column18 { get; set; }
        public bool Column19 { get; set; }
        public bool Column20 { get; set; }
        public bool Column21 { get; set; }
        public bool Column22 { get; set; }
        public bool Column23 { get; set; }
        public bool Column24 { get; set; }
        public bool Column25 { get; set; }
        public bool Column26 { get; set; }
        public bool Column27 { get; set; }
        public bool Column28 { get; set; }
        public bool Column29 { get; set; }
        public bool Column30 { get; set; }
        public bool Column31 { get; set; }
        public bool Column32 { get; set; }
        public bool Column33 { get; set; }
        public bool Column34 { get; set; }
        public bool Column35 { get; set; }
        public ChosenColumnsViewModel() { }
        public ChosenColumnsViewModel(ChosenColumnsDto Item)
        {
            UserID = Item.UserID;
            TableID = Item.TableID;
            Column1 = Item.Column1;
            Column2 = Item.Column2;
            Column3 = Item.Column3;
            Column4 = Item.Column4;
            Column5 = Item.Column5;
            Column6 = Item.Column6;
            Column7 = Item.Column7;
            Column8 = Item.Column8;
            Column9 = Item.Column9;
            Column10 = Item.Column10;
            Column11 = Item.Column12;
            Column13 = Item.Column13;
            Column14 = Item.Column14;
            Column15 = Item.Column15;
        }

        //This Is For When User Wants To Turn Back From RevisionsEditChosenColumns To The Same Revisions Page
        public string FormType { get; set; }
        public string ProjectCode { get; set; }
        public int TaskID { get; set; }
        public string DocumentNumber { get; set; }
        public string DocumentTitle { get; set; }
    }
}
