using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Abp.Application.Services.Dto;
using Abp.AspNetCore.Mvc.Authorization;
using Abp.Runtime.Validation;
using Abp.Web.Models;
using Dapna.MSVPortal.Authorization.Users;
using Dapna.MSVPortal.Controllers;
using Dapna.MSVPortal.Enums;
using Dapna.MSVPortal.Financial;
using Dapna.MSVPortal.Financial.Dto;
using Dapna.MSVPortal.Projects;
using Dapna.MSVPortal.Projects.Dto;
using Dapna.MSVPortal.Web.ViewModels;
using Dapna.MSVPortal.Web.ViewModels.ExcelModel;
using ExcelDataReader;
using Kendo.Mvc.Extensions;
using Kendo.Mvc.UI;
using MD.PersianDateTime;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Rendering;
using Newtonsoft.Json;
using Newtonsoft.Json.Serialization;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using Slack.Client;

namespace Dapna.MSVPortal.Web.Mvc.Controllers
{
    [AbpMvcAuthorize]
    public class FinancialController : MSVPortalControllerBase
    {
        private readonly IHostingEnvironment _hostingEnvironment;
        private readonly IPaymentRequestAppService _paymentRequestAppService;
        private readonly IReceiveBillAppService _receiveBillAppService;
        private readonly ITransactionAppService _transactionAppService;
        private readonly IProjectAppService _projectAppService;
        private readonly IContractorAppService _contractorAppService;
        private readonly IKatibe_DarkhastAppService _katibe_DarkhastAppService;
        private readonly IProject_User_MappingAppService _project_User_MappingAppService;
        private readonly IChosenColumnsAppService _chosenColumnsAppService;
        private readonly IRemainCreditAppService _remainCreditAppService;
        private readonly UserManager _userManager;

        public FinancialController(IPaymentRequestAppService paymentRequestAppService,
            IHostingEnvironment hostingEnvironment,
            IReceiveBillAppService receiveBillAppService,
            ITransactionAppService transactionAppService,
            IProjectAppService projectAppService,
            IContractorAppService contractorAppService,
            IKatibe_DarkhastAppService katibe_DarkhastAppService,
            IProject_User_MappingAppService project_User_MappingAppService,
            IChosenColumnsAppService chosenColumnsAppService,
            IRemainCreditAppService remainCreditAppService,
            UserManager userManager
            )
        {
            _hostingEnvironment = hostingEnvironment;
            _paymentRequestAppService = paymentRequestAppService;
            _receiveBillAppService = receiveBillAppService;
            _transactionAppService = transactionAppService;
            _projectAppService = projectAppService;
            _contractorAppService = contractorAppService;
            _katibe_DarkhastAppService = katibe_DarkhastAppService;
            _project_User_MappingAppService = project_User_MappingAppService;
            _chosenColumnsAppService = chosenColumnsAppService;
            _remainCreditAppService = remainCreditAppService;
            _userManager = userManager;
        }

        public IActionResult ImportData()
        {
            return View();
        }

        //[HttpPost]
        //public async Task<ActionResult> ImportData(IFormFile fileExcel)
        //{
        //    if (Request != null)
        //    {
        //        if ((fileExcel != null) && (fileExcel.Length != 0) && !string.IsNullOrEmpty(fileExcel.FileName))
        //        {
        //            var fileName = Path.GetFileName(fileExcel.FileName);
        //            var newFileName = "msv_" + Guid.NewGuid().ToString("N") + "_ExcelTemp" + "." + fileName.Split('.').LastOrDefault();
        //            try
        //            {
        //                var physicalPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot/Uploads", newFileName);
        //                using (var stream = new FileStream(physicalPath, FileMode.Create))
        //                {
        //                    await fileExcel.CopyToAsync(stream);
        //                }
        //            }
        //            catch (Exception ex)
        //            {
        //                //model.ErrorText = "خطا در زمان ذخیره فایل.";
        //                //_logService.Save(string.Format("Can not save file : {0}/tIndex. {1}", fileName, ex));
        //                return View();
        //            }

        //            var itemsList = new List<ExcelRowViewModel>();


        //            using (var package = new ExcelPackage(new FileInfo(Path.Combine(Directory.GetCurrentDirectory(), "wwwroot/Uploads", newFileName))))
        //            {
        //                var currentSheet = package.Workbook.Worksheets;
        //                var workSheet = currentSheet.First();
        //                var noOfCol = workSheet.Dimension.End.Column;
        //                var noOfRow = workSheet.Dimension.End.Row;
        //                var i = 1;
        //                for (int rowIterator = 1; rowIterator <= noOfRow; rowIterator++)
        //                {
        //                    try
        //                    {
        //                        var row = new ExcelRowViewModel();
        //                        if (workSheet.Cells[rowIterator, 3].Value == null)
        //                        {
        //                            continue;
        //                        }
        //                        row.ProjectId = Convert.ToInt32(workSheet.Cells[rowIterator, 3].Value);
        //                        row.CreationTime = PersianDateTime.Parse(workSheet.Cells[rowIterator, 4].Value.ToString());
        //                        row.FourthLevelCode = workSheet.Cells[rowIterator, 5].Value != null ? workSheet.Cells[rowIterator, 5].Value.ToString() : null;
        //                        row.PaymentRequestCode = workSheet.Cells[rowIterator, 6].Value != null ? workSheet.Cells[rowIterator, 6].Value.ToString() : null;
        //                        row.Description = workSheet.Cells[rowIterator, 7].Value != null ? workSheet.Cells[rowIterator, 7].Value.ToString() : null;
        //                        row.PaymentRequestType = Convert.ToInt32(workSheet.Cells[rowIterator, 9].Value);
        //                        row.PayTo = workSheet.Cells[rowIterator, 10].Value != null ? workSheet.Cells[rowIterator, 10].Value.ToString() : null;
        //                        row.Amount = decimal.Parse(workSheet.Cells[rowIterator, 11].Value.ToString());
        //                        row.PaymentType = Convert.ToInt32(workSheet.Cells[rowIterator, 12].Value);
        //                        row.ChequeDueDate = workSheet.Cells[rowIterator, 14].Value != null ? PersianDateTime.Parse(workSheet.Cells[rowIterator, 14].Value.ToString()) : default(DateTime?);
        //                        row.CurrencyType = Convert.ToInt32(workSheet.Cells[rowIterator, 16].Value);
        //                        row.MoreDescription = workSheet.Cells[rowIterator, 17].Value != null ? workSheet.Cells[rowIterator, 17].Value.ToString() : null;
        //                        itemsList.Add(row);
        //                        i++;
        //                    }
        //                    catch (Exception ex)
        //                    {
        //                        ViewBag.Error = i + " - "
        //                            + ex.Message;
        //                        return View();
        //                    }

        //                }


        //                foreach (var item in itemsList)
        //                {
        //                    var paymentRequestDto = new PaymentRequestDto()
        //                    {
        //                        ProjectId = item.ProjectId,
        //                        CreationTime = item.CreationTime,
        //                        FourthLevelCode = item.FourthLevelCode,
        //                        PaymentRequestCode = item.PaymentRequestCode,
        //                        Description = item.Description,
        //                        PaymentRequestType = (Enums.PaymentRequestType)item.PaymentRequestType,
        //                        PayTo = item.PayTo,
        //                        Amount = item.Amount,
        //                        PaymentType = (Enums.PaymentType)item.PaymentType,
        //                        ChequeDueDate = item.ChequeDueDate,
        //                        CurrencyType = (Enums.CurrencyType)item.CurrencyType,
        //                    };

        //                    var paymentResult = await _paymentRequestAppService.Create(paymentRequestDto);

        //                    var transactionDto = new TransactionDto()
        //                    {
        //                        TransactionType = Enums.TransactionType.Payment,
        //                        Amount = paymentResult.Amount,
        //                        CreationTime = paymentResult.CreationTime,
        //                        CurrencyType = paymentResult.CurrencyType,
        //                        PaymentType = paymentResult.PaymentType,
        //                        Description = item.MoreDescription,
        //                        PaymentRequestId = paymentResult.Id,
        //                        ProjectId = paymentResult.ProjectId,
        //                        ChequeDueDate = paymentResult.ChequeDueDate
        //                    };

        //                    var transResult = await _transactionAppService.Create(transactionDto);

        //                }


        //            }
        //            return RedirectToAction("PaymentRequests");
        //        }
        //    }
        //    return RedirectToAction("PaymentRequests");
        //}


        [HttpPost]
        public async Task<ActionResult> ImportData(IFormFile fileExcel)
        {
            if (Request != null)
            {
                if ((fileExcel != null) && (fileExcel.Length != 0) && !string.IsNullOrEmpty(fileExcel.FileName))
                {
                    var fileName = Path.GetFileName(fileExcel.FileName);
                    var newFileName = "msv_" + Guid.NewGuid().ToString("N") + "_ExcelTemp" + "." + fileName.Split('.').LastOrDefault();
                    try
                    {
                        var physicalPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot/Uploads", newFileName);
                        using (var stream = new FileStream(physicalPath, FileMode.Create))
                        {
                            await fileExcel.CopyToAsync(stream);
                        }
                    }
                    catch (Exception ex)
                    {
                        //model.ErrorText = "خطا در زمان ذخیره فایل.";
                        //_logService.Save(string.Format("Can not save file : {0}/tIndex. {1}", fileName, ex));
                        return View();
                    }

                    var itemsList = new List<ExcelRowViewModel>();


                    using (var package = new ExcelPackage(new FileInfo(Path.Combine(Directory.GetCurrentDirectory(), "wwwroot/Uploads", newFileName))))
                    {
                        var currentSheet = package.Workbook.Worksheets;
                        var workSheet = currentSheet.First();
                        var noOfCol = workSheet.Dimension.End.Column;
                        var noOfRow = workSheet.Dimension.End.Row;
                        var i = 1;
                        for (int rowIterator = 1; rowIterator <= noOfRow; rowIterator++)
                        {
                            if (workSheet.Cells[rowIterator, 1].Value == null)
                            {
                                continue;
                            }
                            try
                            {
                                var row = new ExcelRowViewModel();
                                row.Description = workSheet.Cells[rowIterator, 1].Value.ToString();
                                row.FourthLevelCode = workSheet.Cells[rowIterator, 2].Value != null ? workSheet.Cells[rowIterator, 2].Value.ToString() : null;
                                itemsList.Add(row);
                                i++;
                            }
                            catch (Exception ex)
                            {
                                ViewBag.Error = i + " - "
                                    + ex.Message;
                                return View();
                            }

                        }


                        foreach (var item in itemsList)
                        {
                            var paymentRequestDto = new ContractorDto()
                            {
                                Name = item.FourthLevelCode,
                                Code = item.Description
                            };

                            var paymentResult = await _contractorAppService.Create(paymentRequestDto);

                        }


                    }
                    return RedirectToAction("PaymentRequests");
                }
            }
            return RedirectToAction("PaymentRequests");
        }




        public IActionResult ImportReceipt()
        {
            return View();
        }

        [HttpPost]
        public async Task<ActionResult> ImportReceipt(IFormFile fileExcel)
        {
            if (Request != null)
            {
                if ((fileExcel != null) && (fileExcel.Length != 0) && !string.IsNullOrEmpty(fileExcel.FileName))
                {
                    var fileName = Path.GetFileName(fileExcel.FileName);
                    var newFileName = "msv_" + Guid.NewGuid().ToString("N") + "_ExcelTemp" + "." + fileName.Split('.').LastOrDefault();
                    try
                    {
                        var physicalPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot/Uploads", newFileName);
                        using (var stream = new FileStream(physicalPath, FileMode.Create))
                        {
                            await fileExcel.CopyToAsync(stream);
                        }
                    }
                    catch (Exception ex)
                    {
                        //model.ErrorText = "خطا در زمان ذخیره فایل.";
                        //_logService.Save(string.Format("Can not save file : {0}/tIndex. {1}", fileName, ex));
                        return View();
                    }

                    var itemsList = new List<ExcelRecieptViewModel>();


                    using (var package = new ExcelPackage(new FileInfo(Path.Combine(Directory.GetCurrentDirectory(), "wwwroot/Uploads", newFileName))))
                    {
                        var currentSheet = package.Workbook.Worksheets;
                        var workSheet1 = currentSheet.First();
                        var noOfCol = workSheet1.Dimension.End.Column;
                        var noOfRow = workSheet1.Dimension.End.Row;
                        var i = 1;
                        for (int rowIterator = 1; rowIterator <= noOfRow; rowIterator++)
                        {
                            try
                            {
                                var row = new ExcelRecieptViewModel();
                                if (workSheet1.Cells[rowIterator, 3].Value == null)
                                {
                                    continue;
                                }
                                row.Key = Convert.ToInt32(workSheet1.Cells[rowIterator, 1].Value);
                                row.ProjectId = Convert.ToInt32(workSheet1.Cells[rowIterator, 3].Value);
                                row.CreationTime = PersianDateTime.Parse(workSheet1.Cells[rowIterator, 10].Value.ToString());
                                row.Amount = decimal.Parse(workSheet1.Cells[rowIterator, 7].Value.ToString());
                                row.PaymentType = Convert.ToInt32(workSheet1.Cells[rowIterator, 5].Value);
                                row.CurrencyType = Convert.ToInt32(workSheet1.Cells[rowIterator, 9].Value);
                                itemsList.Add(row);
                                i++;
                            }
                            catch (Exception ex)
                            {
                                ViewBag.Error = i + " -i "
                                    + ex.Message;
                                return View();
                            }

                        }


                        var transList = new List<ExcelRecieptTransactionViewModel>();


                        var workSheet2 = currentSheet.Last();
                        var noOfCol1 = workSheet2.Dimension.End.Column;
                        var noOfRow2 = workSheet2.Dimension.End.Row;
                        var m = 1;
                        for (int rowIterator = 1; rowIterator <= noOfRow2; rowIterator++)
                        {
                            try
                            {
                                var row = new ExcelRecieptTransactionViewModel();
                                if (workSheet2.Cells[rowIterator, 3].Value == null)
                                {
                                    continue;
                                }
                                row.ParentKey = Convert.ToInt32(workSheet2.Cells[rowIterator, 1].Value);
                                row.ProjectId = Convert.ToInt32(workSheet2.Cells[rowIterator, 3].Value);
                                row.CreationTime = PersianDateTime.Parse(workSheet2.Cells[rowIterator, 5].Value.ToString());
                                row.Amount = decimal.Parse(workSheet2.Cells[rowIterator, 4].Value.ToString());
                                row.CurrencyType = 1;
                                transList.Add(row);
                                m++;
                            }
                            catch (Exception ex)
                            {
                                ViewBag.Error = m + " -m "
                                    + ex.Message;
                                return View();
                            }

                        }


                        foreach (var item in itemsList)
                        {
                            var receiveBillDto = new ReceiveBillDto()
                            {
                                ProjectId = item.ProjectId,
                                CreationTime = item.CreationTime,
                                Amount = item.Amount,
                                PaymentType = Enums.PaymentType.Cash,
                                CurrencyType = Enums.CurrencyType.IRR,
                            };

                            var receiveBillDtoResult = await _receiveBillAppService.Create(receiveBillDto);

                            var transitems = transList.Where(a => a.ParentKey == item.Key).ToList();

                            foreach (var trans in transitems)
                            {
                                var transactionDto = new TransactionDto()
                                {
                                    TransactionType = Enums.TransactionType.Receive,
                                    Amount = trans.Amount,
                                    CreationTime = trans.CreationTime,
                                    CurrencyType = Enums.CurrencyType.IRR,
                                    PaymentType = Enums.PaymentType.Cash,
                                    ProjectId = trans.ProjectId,
                                    ReceiveBillId = receiveBillDtoResult.Id
                                };

                                await _transactionAppService.Create(transactionDto);
                            }


                        }


                    }
                    return RedirectToAction("PaymentRequests");
                }
            }
            return RedirectToAction("PaymentRequests");
        }

        #region ChosenColumns
        [HttpPost]
        [DontWrapResult]
        [DisableValidation]
        //Check
        public async Task<IActionResult> AddOrUpdateChosenColumns(ChosenColumnsViewModel Model)
        {
            //GetChosenColumns Returns All The Users Tables , And Then We Sholud Check If The Tables ID Exists
            var Record = _chosenColumnsAppService.GetChosenColumns(AbpSession.UserId.Value).Result.Find(e => e.TableID == Model.TableID);
            if (Record == null)
            {
                string TableSection = "";

                if (Model.TableID == 1)
                {
                    TableSection = "/Views/Financial/PaymentRequests";
                }
                else if (Model.TableID == 2)
                {
                    TableSection = "/Views/Financial/ReceiveBills";
                }
                else if (Model.TableID == 3)
                {
                    TableSection = "/Views/Financial/Transactions";
                }
                else if (Model.TableID == 4)
                {
                    TableSection = "/Views/Financial/QueueOnePaymentRequests";
                }
                else if (Model.TableID == 5)
                {
                    TableSection = "/Views/Financial/QueueTwoPaymentRequests";
                }

                await _chosenColumnsAppService.Create(new ChosenColumnsDto()
                {
                    UserID = (int)AbpSession.UserId.Value,
                    Section = TableSection,
                    TableID = Model.TableID,
                    Column1 = Model.Column1,
                    Column2 = Model.Column2,
                    Column3 = Model.Column3,
                    Column4 = Model.Column4,
                    Column5 = Model.Column5,
                    Column6 = Model.Column6,
                    Column7 = Model.Column7,
                    Column8 = Model.Column8,
                    Column9 = Model.Column9,
                    Column10 = Model.Column10,
                    Column11 = Model.Column11,
                    Column12 = Model.Column12,
                    Column13 = Model.Column13,
                    Column14 = Model.Column14,
                    Column15 = Model.Column15,
                });
            }
            else
            {
                await _chosenColumnsAppService.Update(new ChosenColumnsDto
                {
                    Id = Record.Id,
                    UserID = Record.UserID,
                    TableID = Record.TableID,
                    Column1 = Model.Column1,
                    Column2 = Model.Column2,
                    Column3 = Model.Column3,
                    Column4 = Model.Column4,
                    Column5 = Model.Column5,
                    Column6 = Model.Column6,
                    Column7 = Model.Column7,
                    Column8 = Model.Column8,
                    Column9 = Model.Column9,
                    Column10 = Model.Column10,
                    Column11 = Model.Column11,
                    Column12 = Model.Column12,
                    Column13 = Model.Column13,
                    Column14 = Model.Column14,
                    Column15 = Model.Column15,
                });
            }

            if (Model.TableID == 1)
            {
                return RedirectToAction("PaymentRequests");
            }
            else if (Model.TableID == 2)
            {
                return RedirectToAction("ReceiveBills");
            }
            else if (Model.TableID == 3)
            {
                return RedirectToAction("Transactions");
            }
            else if (Model.TableID == 4)
            {
                return RedirectToAction("QueueOnePaymentRequests");
            }
            else if (Model.TableID == 5)
            {
                return RedirectToAction("QueueTwoPaymentRequests");
            }

            return RedirectToAction("PaymentRequests");
        }
        #endregion


        public async Task<IActionResult> PaymentRequests(int? lastPage = 1)
        {
            var userId = AbpSession.UserId.Value;
            var items = await _project_User_MappingAppService.GetUserProjects(userId);
            ViewBag.ProjectId = items.Select(a => new ProjectViewModel()
            {
                Id = a.ProjectId,
                Title = a.Project.Title
            }).ToList();
            ViewBag.LastPageValue = lastPage;

            //Users ChosenColumns For PaymentRequestTable => PayementRequestTable's ID Equals To 1 => TableID = 1 ;
            var ChosenColumns = _chosenColumnsAppService.GetChosenColumns(AbpSession.UserId.Value).Result.Find(e=>e.TableID==1);

            if (ChosenColumns != null)
            {
                ViewBag.UserID = ChosenColumns.UserID;
                ViewBag.TableID = ChosenColumns.TableID;
                ViewBag.Column1 = ChosenColumns.Column1;
                ViewBag.Column2 = ChosenColumns.Column2;
                ViewBag.Column3 = ChosenColumns.Column3;
                ViewBag.Column4 = ChosenColumns.Column4;
                ViewBag.Column5 = ChosenColumns.Column5;
                ViewBag.Column6 = ChosenColumns.Column6;
                ViewBag.Column7 = ChosenColumns.Column7;
                ViewBag.Column8 = ChosenColumns.Column8;
                ViewBag.Column9 = ChosenColumns.Column9;
                ViewBag.Column10 = ChosenColumns.Column10;
                ViewBag.Column11 = ChosenColumns.Column11;
                ViewBag.Column12 = ChosenColumns.Column12;
                ViewBag.Column13 = ChosenColumns.Column13;
                ViewBag.Column14 = ChosenColumns.Column14;
                ViewBag.Column15 = ChosenColumns.Column15;

                return View();
            }

            return View();
        }

        //Before EditChosenColumns Page Opens , If Database Was Filled For This User & Table , It Fills The ChosenColumnViewModel Before Page Loads Then Returns ChosenColumnsViewModel To EditChosenColumns Page In Order To Edit Them
        public IActionResult PaymentRequests_EditChosenColumns(int TableID)
        {
            var Record = _chosenColumnsAppService.GetChosenColumns(AbpSession.UserId.Value).Result.Find(e => e.TableID == TableID);

            if (Record != null)
            {
                ChosenColumnsViewModel ChosenColumnsViewModel = new ChosenColumnsViewModel
                {
                    UserID = Record.UserID,
                    TableID = Record.TableID,
                    Column1 = Record.Column1,
                    Column2 = Record.Column2,
                    Column3 = Record.Column3,
                    Column4 = Record.Column4,
                    Column5 = Record.Column5,
                    Column6 = Record.Column6,
                    Column7 = Record.Column7,
                    Column8 = Record.Column8,
                    Column9 = Record.Column9,
                    Column10 = Record.Column10,
                    Column11 = Record.Column11,
                    Column12 = Record.Column12,
                    Column13 = Record.Column13,
                    Column14 = Record.Column14,
                    Column15 = Record.Column15
                };
                return View(ChosenColumnsViewModel);
            }
            return View();
        }

        public async Task<IActionResult> QueueOnePaymentRequests(int? lastPage = 1)
        {
            var userId = AbpSession.UserId.Value;
            var items = await _project_User_MappingAppService.GetUserProjects(userId);
            ViewBag.ProjectId = items.Select(a => new ProjectViewModel()
            {
                Id = a.ProjectId,
                Title = a.Project.Title
            }).ToList();

            ViewBag.LastPageValue = lastPage;

            //Users ChosenColumns For PaymentRequestTable => PayementRequestTable's ID Equals To 1 => TableID = 4 ;
            var ChosenColumns = _chosenColumnsAppService.GetChosenColumns(AbpSession.UserId.Value).Result.Find(e => e.TableID == 4);

            if (ChosenColumns != null)
            {
                ViewBag.UserID = ChosenColumns.UserID;
                ViewBag.TableID = ChosenColumns.TableID;
                ViewBag.Column1 = ChosenColumns.Column1;
                ViewBag.Column2 = ChosenColumns.Column2;
                ViewBag.Column3 = ChosenColumns.Column3;
                ViewBag.Column4 = ChosenColumns.Column4;
                ViewBag.Column5 = ChosenColumns.Column5;
                ViewBag.Column6 = ChosenColumns.Column6;
                ViewBag.Column7 = ChosenColumns.Column7;
                ViewBag.Column8 = ChosenColumns.Column8;
                ViewBag.Column9 = ChosenColumns.Column9;
                ViewBag.Column10 = ChosenColumns.Column10;
                ViewBag.Column11 = ChosenColumns.Column11;
                ViewBag.Column12 = ChosenColumns.Column12;
                ViewBag.Column13 = ChosenColumns.Column13;
                ViewBag.Column14 = ChosenColumns.Column14;
                ViewBag.Column15 = ChosenColumns.Column15;

                return View();
            }

            return View();
        }

        public IActionResult QueueOne_EditChosenColumns(int TableID)
        {
            var Record = _chosenColumnsAppService.GetChosenColumns(AbpSession.UserId.Value).Result.Find(e => e.TableID == TableID);

            if (Record != null)
            {
                ChosenColumnsViewModel ChosenColumnsViewModel = new ChosenColumnsViewModel
                {
                    UserID = Record.UserID,
                    TableID = Record.TableID,
                    Column1 = Record.Column1,
                    Column2 = Record.Column2,
                    Column3 = Record.Column3,
                    Column4 = Record.Column4,
                    Column5 = Record.Column5,
                    Column6 = Record.Column6,
                    Column7 = Record.Column7,
                    Column8 = Record.Column8,
                    Column9 = Record.Column9,
                    Column10 = Record.Column10,
                    Column11 = Record.Column11,
                    Column12 = Record.Column12,
                    Column13 = Record.Column13,
                    Column14 = Record.Column14,
                    Column15 = Record.Column15
                };
                return View(ChosenColumnsViewModel);
            }
            return View();
        }

        public async Task<IActionResult> QueueTwoPaymentRequests(int? lastPage = 1)
        {
            var userId = AbpSession.UserId.Value;
            var items = await _project_User_MappingAppService.GetUserProjects(userId);
            ViewBag.ProjectId = items.Select(a => new ProjectViewModel()
            {
                Id = a.ProjectId,
                Title = a.Project.Title
            }).ToList();

            ViewBag.LastPageValue = lastPage;

            //Users ChosenColumns For PaymentRequestTable => PayementRequestTable's ID Equals To 1 => TableID = 5 ;
            var ChosenColumns = _chosenColumnsAppService.GetChosenColumns(AbpSession.UserId.Value).Result.Find(e => e.TableID == 5);

            if (ChosenColumns != null)
            {
                ViewBag.UserID = ChosenColumns.UserID;
                ViewBag.TableID = ChosenColumns.TableID;
                ViewBag.Column1 = ChosenColumns.Column1;
                ViewBag.Column2 = ChosenColumns.Column2;
                ViewBag.Column3 = ChosenColumns.Column3;
                ViewBag.Column4 = ChosenColumns.Column4;
                ViewBag.Column5 = ChosenColumns.Column5;
                ViewBag.Column6 = ChosenColumns.Column6;
                ViewBag.Column7 = ChosenColumns.Column7;
                ViewBag.Column8 = ChosenColumns.Column8;
                ViewBag.Column9 = ChosenColumns.Column9;
                ViewBag.Column10 = ChosenColumns.Column10;
                ViewBag.Column11 = ChosenColumns.Column11;
                ViewBag.Column12 = ChosenColumns.Column12;
                ViewBag.Column13 = ChosenColumns.Column13;
                ViewBag.Column14 = ChosenColumns.Column14;
                ViewBag.Column15 = ChosenColumns.Column15;

                return View();
            }

            return View();
        }

        public IActionResult QueueTwo_EditChosenColumns(int TableID)
        {
            var Record = _chosenColumnsAppService.GetChosenColumns(AbpSession.UserId.Value).Result.Find(e => e.TableID == TableID);

            if (Record != null)
            {
                ChosenColumnsViewModel ChosenColumnsViewModel = new ChosenColumnsViewModel
                {
                    UserID = Record.UserID,
                    TableID = Record.TableID,
                    Column1 = Record.Column1,
                    Column2 = Record.Column2,
                    Column3 = Record.Column3,
                    Column4 = Record.Column4,
                    Column5 = Record.Column5,
                    Column6 = Record.Column6,
                    Column7 = Record.Column7,
                    Column8 = Record.Column8,
                    Column9 = Record.Column9,
                    Column10 = Record.Column10,
                    Column11 = Record.Column11,
                    Column12 = Record.Column12,
                    Column13 = Record.Column13,
                    Column14 = Record.Column14,
                    Column15 = Record.Column15
                };
                return View(ChosenColumnsViewModel);
            }
            return View();
        }

        public async Task<IActionResult> DeclinedPaymentRequests(int? lastPage = 1)
        {
            var userId = AbpSession.UserId.Value;
            var items = await _project_User_MappingAppService.GetUserProjects(userId);
            ViewBag.ProjectId = items.Select(a => new ProjectViewModel()
            {
                Id = a.ProjectId,
                Title = a.Project.Title
            }).ToList();

            ViewBag.LastPageValue = lastPage;

            return View();
        }

        public async Task<IActionResult> AddPaymentRequest()
        {
            var userId = AbpSession.UserId.Value;
            var items = await _project_User_MappingAppService.GetUserProjects(userId);
            ViewBag.ProjectId = items.Select(a => new ProjectViewModel()
            {
                Id = a.ProjectId,
                Title = a.Project.Title
            }).ToList();
            return View();
        }

        [HttpPost]
        public async Task<IActionResult> AddPaymentRequest(PaymentRequestViewModel model)
        {
            var item = new PaymentRequestDto()
            {
                ScheduleRow = model.ScheduleRow,
                FourthLevelCode = model.FourthLevelCode,
                PaymentRequestCode = model.PaymentRequestCode,
                PaymentDescription = model.PaymentDescription,
                PaymentRequestType = model.PaymentRequestType,
                PayTo = model.PayTo,
                Amount = decimal.Parse(model.Amount),
                CreditBalance = model.CreditBalance != null ? decimal.Parse(model.CreditBalance) : default(decimal?),
                CurrencyType = model.CurrencyType,
                PaymentType = model.PaymentType,
                ChequeDueDate = model.ChequeDueDate != null ? PersianDateTime.Parse(model.ChequeDueDate) : default(DateTime?),
                Description = model.Description,
                ProjectId = model.ProjectId,
                PaymentRequestStatus = PaymentRequestStatus.Waiting
            };
            var result = await _paymentRequestAppService.Create(item);

            return RedirectToAction("PaymentRequests");
        }

        public async Task<IActionResult> EditPaymentRequest(int id, int lastPage, string referer)
        {
            var userId = AbpSession.UserId.Value;
            var items = await _project_User_MappingAppService.GetUserProjects(userId);
            ViewBag.ProjectId = items.Select(a => new ProjectViewModel()
            {
                Id = a.ProjectId,
                Title = a.Project.Title
            }).ToList();
            var item = await _paymentRequestAppService.GetWillDetails(id);

            var itemDto = ObjectMapper.Map<PaymentRequestDto>(item);

            var model = new PaymentRequestViewModel(itemDto);
            model.LastGridPage = lastPage;
            model.RefererPage = referer;


            return View(model);
        }

        [HttpPost]
        public async Task<IActionResult> EditPaymentRequest(PaymentRequestViewModel model)
        {
            var item = await _paymentRequestAppService.Get(new EntityDto<int>(model.Id));

            item.ScheduleRow = model.ScheduleRow;
            item.FourthLevelCode = model.FourthLevelCode;
            item.PaymentRequestCode = model.PaymentRequestCode;
            item.PaymentDescription = model.PaymentDescription;
            item.PaymentRequestType = model.PaymentRequestType;
            item.PayTo = model.PayTo;
            item.Amount = decimal.Parse(model.Amount);
            item.CreditBalance = model.CreditBalance != null ? decimal.Parse(model.CreditBalance) : default(decimal?);
            item.CurrencyType = model.CurrencyType;
            item.PaymentType = model.PaymentType;
            item.ChequeDueDate = model.ChequeDueDate != null ? PersianDateTime.Parse(model.ChequeDueDate) : default(DateTime?);
            item.Description = model.Description;
            item.ProjectId = model.ProjectId;

            await _paymentRequestAppService.Update(item);
            if (!string.IsNullOrEmpty(model.RefererPage))
            {
                return RedirectToAction(model.RefererPage, new { lastPage = model.LastGridPage });
            }
            return RedirectToAction("PaymentRequests", new { lastPage = model.LastGridPage });

        }

        [DontWrapResult]
        public async Task<ActionResult> PaymentRequestRead([DataSourceRequest] DataSourceRequest request, int paymentRequestStatus, int? id, int[] selectedProjectId,
            int[] currencyId, int[] paymentRequestType, string from, string to, string payto, string paymentRequestCode, string chequeDueDate,
            int[] paymentType, int? sortType, string fourthLevelCode)
        {
            var userId = AbpSession.UserId.Value;
            var projects = await _project_User_MappingAppService.GetUserProjects(userId);
            var projectIds = projects.Select(b => b.ProjectId).ToList();

            var result = await _paymentRequestAppService.GetPaymentRequestsPage(paymentRequestStatus, id, selectedProjectId, currencyId, paymentRequestType, request.Page, request.PageSize,
                from != null ? PersianDateTime.Parse(from) : default(DateTime?),
                to != null ? PersianDateTime.Parse(to) : default(DateTime?),
                chequeDueDate != null ? PersianDateTime.Parse(chequeDueDate) : default(DateTime?),
                paymentRequestCode,
                payto,
                fourthLevelCode,
                projectIds,
                paymentType,
                sortType
                );

            var items = result.Item1.Select(a => new PaymentRequestViewModel()
            {
                Id = a.Id,
                ScheduleRow = a.ScheduleRow,
                FourthLevelCode = a.FourthLevelCode,
                PaymentRequestType = a.PaymentRequestType,
                Amount = a.Amount.ToString("N0"),
                CreditBalance = a.CreditBalance != null ? a.CreditBalance.Value.ToString("N0") : null,
                PaidAmount = a.Payments.Sum(b => b.Amount),
                PaymentsCount = a.Payments.Count(),
                CurrencyType = a.CurrencyType,
                PaymentType = a.PaymentType,
                PayTo = a.PayTo,
                PaymentDescription = a.PaymentDescription,
                ProjectName = a.Project.Title,
                PaymentRequestStatus = (int)a.PaymentRequestStatus,
                PaymentRequestCode = a.PaymentRequestCode,
                ChequeDueDate = a.ChequeDueDate != null ? new PersianDateTime(a.ChequeDueDate).ToShortDateString() : "-",
                CreationDate = new PersianDateTime(a.CreationTime).ToShortDateString()
            });
            var dsResult = new DataSourceResult()
            {
                Data = items,
                Total = result.Item2
            };

            return Json(dsResult, new JsonSerializerSettings() { ContractResolver = new DefaultContractResolver() });
        }



        [DontWrapResult]
        public async Task<ActionResult> PaymentRequestFilterResult(int paymentRequestStatus, int? id, int[] selectedProjectId,
           int[] currencyId, int[] paymentRequestType, string from, string to, string payto, string paymentRequestCode, string chequeDueDate,
           int[] paymentType, int? sortType, string fourthLevelCode)
        {
            var userId = AbpSession.UserId.Value;
            var projects = await _project_User_MappingAppService.GetUserProjects(userId);
            var projectIds = projects.Select(b => b.ProjectId).ToList();

            var result = await _paymentRequestAppService.GetPaymentRequestsPage(paymentRequestStatus, id, selectedProjectId, currencyId, paymentRequestType, null, null,
                from != null ? PersianDateTime.Parse(from) : default(DateTime?),
                to != null ? PersianDateTime.Parse(to) : default(DateTime?),
                chequeDueDate != null ? PersianDateTime.Parse(chequeDueDate) : default(DateTime?),
                paymentRequestCode,
                payto,
                fourthLevelCode,
                projectIds,
                paymentType,
                sortType
                );
            return Json(new
            {
                total = result.Item3.ToString("N0"),
                totalPaid = result.Item4.ToString("N0"),
                totalCredit = result.Item5.ToString("N0")
            }, new JsonSerializerSettings() { ContractResolver = new DefaultContractResolver() });
        }


        public async Task<ActionResult> PaymentRequestExcel(int paymentRequestStatus, int? id, string selectedProjectId,
           string currencyId, string paymentRequestType, string from, string to, string payto, string paymentRequestCode, string chequeDueDate,
           string paymentType, int? sortType, string fourthLevelCode)
        {

            var userId = AbpSession.UserId.Value;
            var projects = await _project_User_MappingAppService.GetUserProjects(userId);
            var projectIds = projects.Select(b => b.ProjectId).ToList();

            int[] selectedProjectId_array = selectedProjectId != null ? selectedProjectId.Split(",").Select(int.Parse).ToArray() : new int[] { };
            int[] currencyId_array = currencyId != null ? currencyId.Split(",").Select(int.Parse).ToArray() : new int[] { };
            int[] paymentRequestType_array = paymentRequestType != null ? paymentRequestType.Split(",").Select(int.Parse).ToArray() : new int[] { };
            int[] paymentType_array = paymentType != null ? paymentType.Split(",").Select(int.Parse).ToArray() : new int[] { };


            var result = await _paymentRequestAppService.GetPaymentRequestsPage(paymentRequestStatus, id, selectedProjectId_array, currencyId_array, paymentRequestType_array, null, null,
                         from != null ? PersianDateTime.Parse(from) : default(DateTime?),
                         to != null ? PersianDateTime.Parse(to) : default(DateTime?),
                         chequeDueDate != null ? PersianDateTime.Parse(chequeDueDate) : default(DateTime?),
                         paymentRequestCode,
                         payto,
                         fourthLevelCode,
                         projectIds,
                         paymentType_array,
                         sortType
                         );


            var items = result.Item1;
            var total = result.Item2;

            var fileName = "MSVCO-PaymentRequests-" + DateTime.Now.ToString("yyyy-MM-dd--hh-mm-ss") + ".xlsx";
            var outputDir = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot/Uploads");

            // Create the file using the FileInfo object
            var file = new FileInfo(outputDir + fileName);
            using (var package = new ExcelPackage(file))
            {
                // add a new worksheet to the empty workbook
                ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("MSVCO Payment Requests List - " + DateTime.Now.ToShortDateString());
                worksheet.HeaderFooter.FirstFooter.LeftAlignedText = string.Format("Generated: {0}", DateTime.Now.ToShortDateString());
                worksheet.Row(1).Height = 15;
                worksheet.View.RightToLeft = true;

                // Start adding the header
                worksheet.Cells[1, 1].Value = "شناسه";
                worksheet.Cells[1, 2].Value = "پروژه";
                worksheet.Cells[1, 3].Value = "کد پروژه";
                worksheet.Cells[1, 4].Value = "تاریخ ثبت";
                worksheet.Cells[1, 5].Value = "تاریخ آخرین تغییر";
                worksheet.Cells[1, 6].Value = "کد سطح چهارم	";
                worksheet.Cells[1, 7].Value = "شناسه درخواست وجه";
                worksheet.Cells[1, 8].Value = "شرح";
                worksheet.Cells[1, 9].Value = "طبقه بندی";
                worksheet.Cells[1, 10].Value = "در وجه";
                worksheet.Cells[1, 11].Value = "مانده بستانکار";
                worksheet.Cells[1, 12].Value = "مقدار درخواست";
                worksheet.Cells[1, 13].Value = "تخصیص داده شده";
                worksheet.Cells[1, 14].Value = "واحد پول";
                worksheet.Cells[1, 15].Value = "نحوه پرداخت";
                worksheet.Cells[1, 16].Value = "تاریخ چک	";
                worksheet.Cells[1, 17].Value = "توضیحات درخواست کننده";


                //// Add the second row of header data
                using (var range = worksheet.Cells[1, 1, 1, 17])
                {
                    range.Style.Font.Bold = true;
                    range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    range.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.DarkBlue);
                    range.Style.Font.Color.SetColor(System.Drawing.Color.WhiteSmoke);
                    range.Style.ShrinkToFit = false;
                    range.Style.WrapText = true;
                    range.Style.Locked = true;
                }

                int rowNumber = 2;

                foreach (var item in items)
                {
                    worksheet.Cells[rowNumber, 1].Value = item.Id;
                    worksheet.Cells[rowNumber, 2].Value = item.Project.Title;
                    worksheet.Cells[rowNumber, 3].Value = item.Project.Code;
                    worksheet.Cells[rowNumber, 4].Value = new PersianDateTime(item.CreationTime).ToShortDateString();
                    worksheet.Cells[rowNumber, 5].Value = item.LastModificationTime != null ? new PersianDateTime(item.LastModificationTime).ToShortDateString() : "-";
                    worksheet.Cells[rowNumber, 6].Value = item.FourthLevelCode;
                    worksheet.Cells[rowNumber, 7].Value = item.PaymentRequestCode;
                    worksheet.Cells[rowNumber, 8].Value = item.PaymentDescription;
                    worksheet.Cells[rowNumber, 9].Value = item.PaymentRequestType != null ? item.PaymentRequestType.GetDisplayName() : "-";
                    worksheet.Cells[rowNumber, 10].Value = item.PayTo;
                    worksheet.Cells[rowNumber, 11].Value = item.CreditBalance;
                    worksheet.Cells[rowNumber, 12].Value = item.Amount;
                    worksheet.Cells[rowNumber, 13].Value = item.Payments.Sum(b => b.Amount);
                    worksheet.Cells[rowNumber, 14].Value = item.CurrencyType.GetDisplayName();
                    worksheet.Cells[rowNumber, 15].Value = item.PaymentType.GetDisplayName();
                    worksheet.Cells[rowNumber, 16].Value = item.ChequeDueDate != null ? new PersianDateTime(item.ChequeDueDate).ToShortDateString() : "-";
                    worksheet.Cells[rowNumber, 17].Value = item.Description;

                    rowNumber++;
                }
                //worksheet.Column(1).Width = 5;
                //worksheet.Column(2).Width += 10;
                //worksheet.Column(3).Width += 10;
                //worksheet.Column(4).Width += 70;
                //worksheet.Column(5).Width += 30;
                //worksheet.Column(6).Width += 10;
                //worksheet.Column(7).Width += 10;
                //worksheet.Column(8).Width += 10;
                //worksheet.Column(9).Width += 10;
                //worksheet.Column(10).Width += 10;

                package.Workbook.Properties.Title = "Portal";
                package.Workbook.Properties.Author = "Dapna.Co";
                package.Workbook.Properties.Company = "MSVCO";

                var stream = new MemoryStream();
                package.SaveAs(stream);
                stream.Position = 0;

                const string contentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                stream.Position = 0;
                return File(stream, contentType, fileName);
            }
        }


        [AcceptVerbs("Post")]
        [DontWrapResult]
        [DisableValidation]
        public async Task<IActionResult> PaymentRequestDestroy([DataSourceRequest] DataSourceRequest request, PaymentRequestViewModel item)
        {
            var result = await _paymentRequestAppService.DeletePaymentRequest(item.Id);
            if (!result)
            {
                ModelState.AddModelError("DependentRecord", "true");
                return Json(new[] { item }.ToDataSourceResult(request, ModelState));
            }
            return Json(new[] { item }.ToDataSourceResult(request, ModelState));
        }

        public IActionResult PaymentRequestTransaction(int id, int status)
        {
            return ViewComponent("PaymentRequestTransaction", new { paymentRequestId = id, status });
        }

        public async Task<IActionResult> ReceiveBillTransaction(int id)
        {
            return ViewComponent("ReceiveBillTransaction", new { ReceivebillId = id });
        }

        [HttpPost]
        [DontWrapResult]
        [DisableValidation]
        public async Task<IActionResult> AddPaymentRequestTransaction(string amount, string description, int paymentRequestId, string date)
        {
            try
            {
                var paymentRequest = await _paymentRequestAppService.Get(new EntityDto<int>(paymentRequestId));
                var item = new TransactionDto()
                {
                    TransactionType = Enums.TransactionType.Payment,
                    Amount = Convert.ToDecimal(amount),
                    CurrencyType = paymentRequest.CurrencyType,
                    PaymentType = Enums.PaymentType.Cash,
                    Description = description,
                    PaymentRequestId = paymentRequest.Id,
                    ProjectId = paymentRequest.ProjectId,
                    CreationTime = PersianDateTime.Parse(date)
                };
                paymentRequest.PaymentRequestStatus = PaymentRequestStatus.Confirmed;

                await _paymentRequestAppService.Update(paymentRequest);

                await _transactionAppService.Create(item);

                return Json(new { status = "success" });
            }
            catch (Exception ex)
            {
                return Json(new { status = "error", message = ex.Message });
                throw;
            }
        }


        [DontWrapResult]
        public async Task<ActionResult> PaymentRequestTransactionRead([DataSourceRequest] DataSourceRequest request, int id)
        {
            var items = await _transactionAppService.GetPaymentRequestTransactions(id);
            var paymentrequests = items.Select(a => new TransactionViewModel()
            {
                Id = a.Id,
                Amount = a.Amount,
                CurrencyType = a.CurrencyType,
                Description = a.Description,
                ProjectName = a.Project != null ? a.Project.Title : "-",
                CreationTime = new PersianDateTime(a.CreationTime).ToShortDateString(),
                TransactionType = a.TransactionType
            });
            var dsResult = paymentrequests.ToDataSourceResult(request);
            return Json(dsResult, new JsonSerializerSettings() { ContractResolver = new DefaultContractResolver() });
        }


        [AcceptVerbs("Post")]
        [DontWrapResult]
        [DisableValidation]
        public async Task<IActionResult> PaymentRequestTransactionDestroy([DataSourceRequest] DataSourceRequest request, TransactionViewModel item)
        {
            if (item != null)
            {
                await _transactionAppService.Delete(new EntityDto(item.Id));
            }
            return Json(new[] { item }.ToDataSourceResult(request, ModelState));
        }


        public async Task<IActionResult> ReceiveBills()
        {
            var userId = AbpSession.UserId.Value;
            var items = await _project_User_MappingAppService.GetUserProjects(userId);
            ViewBag.ProjectId = items.Select(a => new ProjectViewModel()
            {
                Id = a.ProjectId,
                Title = a.Project.Title
            }).ToList();

            //Users ChosenColumns For ReceiveBillsTable => ReceiveBillsTable's ID Equals To 2 => TableID = 2 ;
            var ChosenColumns = _chosenColumnsAppService.GetChosenColumns(AbpSession.UserId.Value).Result.Find(e => e.TableID == 2);

            if (ChosenColumns != null)
            {
                ViewBag.UserID = ChosenColumns.UserID;
                ViewBag.TableID = ChosenColumns.TableID;
                ViewBag.Column1 = ChosenColumns.Column1;
                ViewBag.Column2 = ChosenColumns.Column2;
                ViewBag.Column3 = ChosenColumns.Column3;
                ViewBag.Column4 = ChosenColumns.Column4;
                ViewBag.Column5 = ChosenColumns.Column5;
                ViewBag.Column6 = ChosenColumns.Column6;

                return View();
            }

            return View();
        }

        public IActionResult ReceiveBills_EditChosenColumns(int TableID)
        {
            var Record = _chosenColumnsAppService.GetChosenColumns(AbpSession.UserId.Value).Result.Find(e => e.TableID == TableID);

            if (Record != null)
            {
                ChosenColumnsViewModel ChosenColumnsViewModel = new ChosenColumnsViewModel
                {
                    UserID = Record.UserID,
                    TableID = Record.TableID,
                    Column1 = Record.Column1,
                    Column2 = Record.Column2,
                    Column3 = Record.Column3,
                    Column4 = Record.Column4,
                    Column5 = Record.Column5,
                    Column6 = Record.Column6,
                    Column7 = Record.Column7,
                    Column8 = Record.Column8,
                    Column9 = Record.Column9,
                    Column10 = Record.Column10,
                    Column11 = Record.Column11,
                    Column12 = Record.Column12,
                    Column13 = Record.Column13,
                    Column14 = Record.Column14,
                    Column15 = Record.Column15
                };
                return View(ChosenColumnsViewModel);
            }
            return View();
        }


        public async Task<IActionResult> Transactions()
        {
            var userId = AbpSession.UserId.Value;
            var items = await _project_User_MappingAppService.GetUserProjects(userId);
            ViewBag.ProjectId = items.Select(a => new ProjectViewModel()
            {
                Id = a.ProjectId,
                Title = a.Project.Title
            }).ToList();

            //Users ChosenColumns For TransactionsTable => TransactionsTable's ID Equals To 3 => TableID = 3 ;
            var ChosenColumns = _chosenColumnsAppService.GetChosenColumns(AbpSession.UserId.Value).Result.Find(e => e.TableID == 3);

            if (ChosenColumns != null)
            {
                ViewBag.UserID = ChosenColumns.UserID;
                ViewBag.TableID = ChosenColumns.TableID;
                ViewBag.Column1 = ChosenColumns.Column1;
                ViewBag.Column2 = ChosenColumns.Column2;
                ViewBag.Column3 = ChosenColumns.Column3;
                ViewBag.Column4 = ChosenColumns.Column4;
                ViewBag.Column5 = ChosenColumns.Column5;
                ViewBag.Column6 = ChosenColumns.Column6;
                ViewBag.Column7 = ChosenColumns.Column7;
                ViewBag.Column8 = ChosenColumns.Column8;
                ViewBag.Column9 = ChosenColumns.Column9;
                ViewBag.Column10 = ChosenColumns.Column10;
                ViewBag.Column11 = ChosenColumns.Column11;
                ViewBag.Column12 = ChosenColumns.Column12;

                return View();
            }

            return View();
        }
        public IActionResult Transactions_EditChosenColumns(int TableID)
        {
            var Record = _chosenColumnsAppService.GetChosenColumns(AbpSession.UserId.Value).Result.Find(e => e.TableID == TableID);

            if (Record != null)
            {
                ChosenColumnsViewModel ChosenColumnsViewModel = new ChosenColumnsViewModel
                {
                    UserID = Record.UserID,
                    TableID = Record.TableID,
                    Column1 = Record.Column1,
                    Column2 = Record.Column2,
                    Column3 = Record.Column3,
                    Column4 = Record.Column4,
                    Column5 = Record.Column5,
                    Column6 = Record.Column6,
                    Column7 = Record.Column7,
                    Column8 = Record.Column8,
                    Column9 = Record.Column9,
                    Column10 = Record.Column10,
                    Column11 = Record.Column11,
                    Column12 = Record.Column12,
                    Column13 = Record.Column13,
                    Column14 = Record.Column14,
                    Column15 = Record.Column15
                };
                return View(ChosenColumnsViewModel);
            }
            return View();
        }

        [DontWrapResult]
        public async Task<ActionResult> TransactionRead([DataSourceRequest] DataSourceRequest request, int[] projectId, int[] currencyId,
            int[] transactionType, int[] paymentRequestType, string from, string to, string payTo,
            string fourthLeveLCode, string paymentRequestCode, string paymentDescription,
            int? sortType, int[] paymentTypeId)
        {
            var userId = AbpSession.UserId.Value;
            var projects = await _project_User_MappingAppService.GetUserProjects(userId);
            var projectIds = projects.Select(b => b.ProjectId).ToList();

            var result = await _transactionAppService.GetTransactionsPage(projectIds, projectId, currencyId, transactionType, paymentRequestType, request.Page, request.PageSize,
                from != null ? PersianDateTime.Parse(from) : default(DateTime?),
                to != null ? PersianDateTime.Parse(to) : default(DateTime?),
                payTo,
                fourthLeveLCode,
                paymentRequestCode,
                paymentDescription,
                paymentTypeId,
                sortType
                );

            var items = result.Item1.Select(i => new TransactionViewModel
            {
                Id = i.Id,
                Amount = i.Amount,
                CreationTime = new PersianDateTime(i.CreationTime).ToShortDateString(),
                CurrencyType = i.CurrencyType,
                Description = i.Description,
                ProjectName = i.Project != null ? i.Project.Title : "-",
                TransactionType = i.TransactionType,
                PaymentRequest_FourthLevelCode = i.PaymentRequest != null ? i.PaymentRequest.FourthLevelCode : "-",
                PaymentRequest_PayTo = i.PaymentRequest != null ? i.PaymentRequest.PayTo : "-",
                PaymentRequest_Type = i.PaymentRequest != null ? i.PaymentRequest.PaymentRequestType : default(Enums.PaymentRequestType),
                PaymentRequest_PaymentRequestCode = i.PaymentRequest != null ? i.PaymentRequest.PaymentRequestCode : "-",
                PaymentRequest_ChequeDueDate = i.PaymentRequest != null && i.PaymentRequest.ChequeDueDate != null ? new PersianDateTime(i.PaymentRequest.ChequeDueDate).ToShortDateString() : "-",
                PaymentRequest_Description = i.PaymentRequest != null ? i.PaymentRequest.Description : "-",
                PaymentRequest_PaymentDescription = i.PaymentRequest != null ? i.PaymentRequest.PaymentDescription : "-",
                RefrenceId = i.PaymentRequest != null ? i.PaymentRequestId.Value : i.ReceiveBillId.Value,
                //RefrenceDescription = i.PaymentRequest != null ? i.PaymentRequest.PaymentType.ToString() + "درخواست پرداخت " : i.ReceiveBill.ReceiveBillType.ToString() + "سند دریافت "
            }).ToList();

            var dsResult = new DataSourceResult()
            {
                Data = items,
                Total = result.Item2
            };

            return Json(dsResult, new JsonSerializerSettings() { ContractResolver = new DefaultContractResolver() });
        }


        [DontWrapResult]
        public async Task<ActionResult> TransactionFilterResult(int[] projectId, int[] currencyId,
            int[] transactionType, int[] paymentRequestType, string from, string to, string payTo,
            string fourthLeveLCode, string paymentRequestCode, string paymentDescription,
            int? sortType, int[] paymentTypeId)
        {
            var userId = AbpSession.UserId.Value;
            var projects = await _project_User_MappingAppService.GetUserProjects(userId);
            var projectIds = projects.Select(b => b.ProjectId).ToList();

            var result = await _transactionAppService.GetTransactionsPage(projectIds, projectId, currencyId, transactionType, paymentRequestType, null, null,
                from != null ? PersianDateTime.Parse(from) : default(DateTime?),
                to != null ? PersianDateTime.Parse(to) : default(DateTime?),
                payTo,
                fourthLeveLCode,
                paymentRequestCode,
                paymentDescription,
                paymentTypeId,
                sortType
                );

            return Json(result.Item3.ToString("N0"), new JsonSerializerSettings() { ContractResolver = new DefaultContractResolver() });
        }


        [DisableValidation]
        public async Task<ActionResult> TransactionExcel(string projectId, string currencyId,
             string transactionType, string paymentRequestType, string from, string to, string payTo,
            string fourthLeveLCode, string paymentRequestCode, string paymentDescription,
            int? sortType, string paymentTypeId)
        {
            var userId = AbpSession.UserId.Value;
            var projects = await _project_User_MappingAppService.GetUserProjects(userId);
            var projectIds = projects.Select(b => b.ProjectId).ToList();

            int[] projectId_array = projectId != null ? projectId.Split(",").Select(int.Parse).ToArray() : new int[] { };
            int[] currencyId_array = currencyId != null ? currencyId.Split(",").Select(int.Parse).ToArray() : new int[] { };
            int[] transactionType_array = transactionType != null ? transactionType.Split(",").Select(int.Parse).ToArray() : new int[] { };
            int[] paymentRequestType_array = paymentRequestType != null ? paymentRequestType.Split(",").Select(int.Parse).ToArray() : new int[] { };
            int[] paymentTypeId_array = paymentTypeId != null ? paymentTypeId.Split(",").Select(int.Parse).ToArray() : new int[] { };


            var result = await _transactionAppService.GetTransactionsPage(projectIds, projectId_array, currencyId_array, transactionType_array, paymentRequestType_array, null, null,
                from != null ? PersianDateTime.Parse(from) : default(DateTime?),
                to != null ? PersianDateTime.Parse(to) : default(DateTime?),
                payTo,
                fourthLeveLCode,
                paymentRequestCode,
                paymentDescription,
                paymentTypeId_array,
                sortType
                );

            var items = result.Item1;
            var total = result.Item2;

            var fileName = "MSVCO-Transactions-" + DateTime.Now.ToString("yyyy-MM-dd--hh-mm-ss") + ".xlsx";
            var outputDir = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot/Uploads");

            // Create the file using the FileInfo object
            var file = new FileInfo(outputDir + fileName);
            using (var package = new ExcelPackage(file))
            {
                // add a new worksheet to the empty workbook
                ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("MSVCO Transactions List - " + DateTime.Now.ToShortDateString());
                worksheet.HeaderFooter.FirstFooter.LeftAlignedText = string.Format("Generated: {0}", DateTime.Now.ToShortDateString());
                worksheet.Row(1).Height = 15;
                worksheet.View.RightToLeft = true;

                // Start adding the header
                worksheet.Cells[1, 1].Value = "شناسه";
                worksheet.Cells[1, 2].Value = "پروژه";
                worksheet.Cells[1, 3].Value = "کد پروژه";
                worksheet.Cells[1, 4].Value = "تاریخ درخواست";
                worksheet.Cells[1, 5].Value = "تاریخ ثبت";
                worksheet.Cells[1, 6].Value = "عملیات";
                worksheet.Cells[1, 7].Value = "کد سطح چهارم	";
                worksheet.Cells[1, 8].Value = "شناسه درخواست وجه";
                worksheet.Cells[1, 9].Value = "عنوان فعالیت";
                worksheet.Cells[1, 10].Value = "طبقه بندی";
                worksheet.Cells[1, 11].Value = "در وجه";
                worksheet.Cells[1, 12].Value = "مانده بستانکار";
                worksheet.Cells[1, 13].Value = "مقدار درخواستی";
                worksheet.Cells[1, 14].Value = "مقدار";
                worksheet.Cells[1, 15].Value = "واحد پول";
                worksheet.Cells[1, 16].Value = "نحوه پرداخت";
                worksheet.Cells[1, 17].Value = "تاریخ چک";
                worksheet.Cells[1, 18].Value = "توضیحات درخواست کننده";
                worksheet.Cells[1, 19].Value = "توضیحات نهایی";
                worksheet.Cells[1, 20].Value = "شماره سند مرجع";

                //// Add the second row of header data
                using (var range = worksheet.Cells[1, 1, 1, 20])
                {
                    range.Style.Font.Bold = true;
                    range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    range.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.DarkBlue);
                    range.Style.Font.Color.SetColor(System.Drawing.Color.WhiteSmoke);
                    range.Style.ShrinkToFit = false;
                    range.Style.WrapText = true;
                    range.Style.Locked = true;
                }

                int rowNumber = 2;

                foreach (var item in items)
                {
                    worksheet.Cells[rowNumber, 1].Value = item.Id;
                    worksheet.Cells[rowNumber, 2].Value = item.Project.Title;
                    worksheet.Cells[rowNumber, 3].Value = item.Project.Code;
                    worksheet.Cells[rowNumber, 4].Value = item.PaymentRequest != null ? new PersianDateTime(item.PaymentRequest.CreationTime).ToShortDateString() : "-";
                    worksheet.Cells[rowNumber, 5].Value = new PersianDateTime(item.CreationTime).ToShortDateString();
                    worksheet.Cells[rowNumber, 6].Value = item.TransactionType.GetDisplayName();
                    worksheet.Cells[rowNumber, 7].Value = item.PaymentRequest != null ? item.PaymentRequest.FourthLevelCode : "-";
                    worksheet.Cells[rowNumber, 8].Value = item.PaymentRequest != null ? item.PaymentRequest.PaymentRequestCode : "-";
                    worksheet.Cells[rowNumber, 9].Value = item.PaymentRequest != null ? item.PaymentRequest.PaymentDescription : "-";
                    worksheet.Cells[rowNumber, 10].Value = item.PaymentRequest != null && item.PaymentRequest.PaymentRequestType != null ? item.PaymentRequest.PaymentRequestType.GetDisplayName() : "-";
                    worksheet.Cells[rowNumber, 11].Value = item.PaymentRequest != null ? item.PaymentRequest.PayTo : "-";
                    worksheet.Cells[rowNumber, 12].Value = item.PaymentRequest != null ? item.PaymentRequest.CreditBalance : 0;
                    worksheet.Cells[rowNumber, 13].Value = item.PaymentRequest != null ? item.PaymentRequest.Amount : 0;
                    worksheet.Cells[rowNumber, 14].Value = item.Amount;
                    worksheet.Cells[rowNumber, 15].Value = item.CurrencyType.GetDisplayName();
                    worksheet.Cells[rowNumber, 16].Value = item.PaymentRequest != null && item.PaymentRequest.PaymentType != null ? item.PaymentRequest.PaymentType.GetDisplayName() : "-";
                    worksheet.Cells[rowNumber, 17].Value = item.PaymentRequest != null && item.PaymentRequest.ChequeDueDate != null ? new PersianDateTime(item.PaymentRequest.ChequeDueDate).ToShortDateString() : "-";
                    worksheet.Cells[rowNumber, 18].Value = item.PaymentRequest != null ? item.PaymentRequest.Description : "-";
                    worksheet.Cells[rowNumber, 19].Value = item.Description;
                    worksheet.Cells[rowNumber, 20].Value = item.PaymentRequest != null ? item.PaymentRequestId.Value : item.ReceiveBillId.Value;

                    rowNumber++;
                }

                package.Workbook.Properties.Title = "Portal";
                package.Workbook.Properties.Author = "Dapna.Co";
                package.Workbook.Properties.Company = "MSVCO";

                var stream = new MemoryStream();
                package.SaveAs(stream);
                stream.Position = 0;

                const string contentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                stream.Position = 0;
                return File(stream, contentType, fileName);
            }
        }


        public async Task<IActionResult> AddReceiveBill()
        {
            var userId = AbpSession.UserId.Value;
            var items = await _project_User_MappingAppService.GetUserProjects(userId);
            ViewBag.ProjectId = items.Select(a => new ProjectViewModel()
            {
                Id = a.ProjectId,
                Title = a.Project.Title
            }).ToList();
            return View();
        }

        [HttpPost]
        public async Task<IActionResult> AddReceiveBill(ReceiveBillViewModel model)
        {
            var item = new ReceiveBillDto()
            {
                ReceiveBillType = model.ReceiveBillType,
                Amount = decimal.Parse(model.Amount),
                CurrencyType = model.CurrencyType,
                PaymentType = model.PaymentType,
                ReceiveDate = model.ReceiveDate != null ? PersianDateTime.Parse(model.ReceiveDate) : default(DateTime?),
                ChequeDueDate = model.ChequeDueDate != null ? PersianDateTime.Parse(model.ChequeDueDate) : default(DateTime?),
                Description = model.Description,
                ReceiveDescription = model.ReceiveDescription,
                ProjectId = model.ProjectId
            };
            await _receiveBillAppService.Create(item);
            return RedirectToAction("ReceiveBills");
        }


        public async Task<IActionResult> EditReceiveBill(int id)
        {
            var userId = AbpSession.UserId.Value;
            var items = await _project_User_MappingAppService.GetUserProjects(userId);
            ViewBag.ProjectId = items.Select(a => new ProjectViewModel()
            {
                Id = a.ProjectId,
                Title = a.Project.Title
            }).ToList();
            var item = await _receiveBillAppService.Get(new EntityDto<int>(id));
            var model = new ReceiveBillViewModel(item);
            return View(model);
        }

        [HttpPost]
        public async Task<IActionResult> EditReceiveBill(ReceiveBillViewModel model)
        {
            var record = await _receiveBillAppService.Get(new EntityDto<int>(model.Id));

            record.ReceiveBillType = model.ReceiveBillType;
            record.Amount = decimal.Parse(model.Amount);
            record.CurrencyType = model.CurrencyType;
            record.PaymentType = model.PaymentType;
            record.ChequeDueDate = model.ChequeDueDate != null ? PersianDateTime.Parse(model.ChequeDueDate) : default(DateTime?);
            record.ReceiveDate = model.ReceiveDate != null ? PersianDateTime.Parse(model.ReceiveDate) : default(DateTime?);
            record.Description = model.Description;
            record.ReceiveDescription = model.ReceiveDescription;
            record.ProjectId = model.ProjectId;
            await _receiveBillAppService.Update(record);

            return RedirectToAction("ReceiveBills");
        }

        [DontWrapResult]
        public async Task<ActionResult> ReceiveBillsRead([DataSourceRequest] DataSourceRequest request, int? id, int[] projectId, int[] currencyId, int[] receiveBillType, string from, string to, int? sortType)
        {
            var userId = AbpSession.UserId.Value;
            var projects = await _project_User_MappingAppService.GetUserProjects(userId);
            var projectIds = projects.Select(b => b.ProjectId).ToList();

            var result = await _receiveBillAppService.GetReceiveBillsPage(projectIds, id, projectId, currencyId, receiveBillType, request.Page, request.PageSize,
                from != null ? PersianDateTime.Parse(from) : default(DateTime?),
                to != null ? PersianDateTime.Parse(to) : default(DateTime?),
                sortType);

            var items = result.Item1.Select(a => new ReceiveBillViewModel()
            {
                Id = a.Id,
                ReceiveBillType = a.ReceiveBillType,
                Amount = a.Amount.ToString("N0"),
                CurrencyType = a.CurrencyType,
                PaymentType = a.PaymentType,
                ChequeDueDate = a.ChequeDueDate != null ? new PersianDateTime(a.ChequeDueDate).ToShortDateString() : default(string),
                Description = a.Description,
                ReceiveDescription = a.ReceiveDescription,
                ProjectId = a.ProjectId,
                ProjectName = a.Project.Title,
                AssignedAmount = a.Assignemnts.Sum(b => b.Amount),
                Remaining = a.Amount - a.Assignemnts.Sum(b => b.Amount),
                CreationDate = new PersianDateTime(a.CreationTime).ToShortDateString(),
                ReceiveDate = a.ReceiveDate != null ? new PersianDateTime(a.ReceiveDate).ToShortDateString() : default(string),
            });
            var dsResult = new DataSourceResult()
            {
                Data = items,
                Total = result.Item2
            };

            return Json(dsResult, new JsonSerializerSettings() { ContractResolver = new DefaultContractResolver() });
        }

        [DontWrapResult]
        public async Task<ActionResult> ReceiveBillsFilterResult(int? id, int[] projectId, int[] currencyId, int[] receiveBillType, string from, string to, int? sortType)
        {

            var userId = AbpSession.UserId.Value;
            var projects = await _project_User_MappingAppService.GetUserProjects(userId);
            var projectIds = projects.Select(b => b.ProjectId).ToList();

            var result = await _receiveBillAppService.GetReceiveBillsPage(projectIds, id, projectId, currencyId, receiveBillType, null, null,
                from != null ? PersianDateTime.Parse(from) : default(DateTime?),
                to != null ? PersianDateTime.Parse(to) : default(DateTime?),
                sortType);

            return Json(new { total = result.Item3.ToString("N0"), assignedTotal = result.Item4.ToString("N0") }, new JsonSerializerSettings() { ContractResolver = new DefaultContractResolver() });
        }




        public async Task<ActionResult> ReceiveBillsExcel(int? id, string projectId, string currencyId, string receiveBillType, string from, string to, int? sortType)
        {
            var userId = AbpSession.UserId.Value;
            var projects = await _project_User_MappingAppService.GetUserProjects(userId);
            var projectIds = projects.Select(b => b.ProjectId).ToList();

            int[] projectId_array = projectId != null ? projectId.Split(",").Select(int.Parse).ToArray() : new int[] { };
            int[] currencyId_array = currencyId != null ? currencyId.Split(",").Select(int.Parse).ToArray() : new int[] { };
            int[] receiveBillType_array = receiveBillType != null ? receiveBillType.Split(",").Select(int.Parse).ToArray() : new int[] { };


            var result = await _receiveBillAppService.GetReceiveBillsPage(projectIds, id, projectId_array, currencyId_array, receiveBillType_array, null, null,
                from != null ? PersianDateTime.Parse(from) : default(DateTime?),
                to != null ? PersianDateTime.Parse(to) : default(DateTime?),
                sortType);


            var items = result.Item1;
            var total = result.Item2;

            var fileName = "MSVCO-ReceiveBills-" + DateTime.Now.ToString("yyyy-MM-dd--hh-mm-ss") + ".xlsx";
            var outputDir = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot/Uploads");

            // Create the file using the FileInfo object
            var file = new FileInfo(outputDir + fileName);
            using (var package = new ExcelPackage(file))
            {
                // add a new worksheet to the empty workbook
                ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("MSVCO Receive Bills List - " + DateTime.Now.ToShortDateString());
                worksheet.HeaderFooter.FirstFooter.LeftAlignedText = string.Format("Generated: {0}", DateTime.Now.ToShortDateString());
                worksheet.Row(1).Height = 15;
                worksheet.View.RightToLeft = true;

                // Start adding the header
                worksheet.Cells[1, 1].Value = "شناسه";
                worksheet.Cells[1, 2].Value = "پروژه";
                worksheet.Cells[1, 3].Value = "مقدار";
                worksheet.Cells[1, 4].Value = "تخصیص داده شده";
                worksheet.Cells[1, 5].Value = "مانده";
                worksheet.Cells[1, 6].Value = "نحوه دریافت";
                worksheet.Cells[1, 7].Value = "واحد پول	";
                worksheet.Cells[1, 8].Value = "تاریخ";

                //// Add the second row of header data
                using (var range = worksheet.Cells[1, 1, 1, 8])
                {
                    range.Style.Font.Bold = true;
                    range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    range.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.DarkBlue);
                    range.Style.Font.Color.SetColor(System.Drawing.Color.WhiteSmoke);
                    range.Style.ShrinkToFit = false;
                    range.Style.WrapText = true;
                    range.Style.Locked = true;
                }

                int rowNumber = 2;

                foreach (var item in items)
                {
                    worksheet.Cells[rowNumber, 1].Value = item.Id;
                    worksheet.Cells[rowNumber, 2].Value = item.Project.Title;
                    worksheet.Cells[rowNumber, 3].Value = item.Amount;
                    worksheet.Cells[rowNumber, 4].Value = item.Assignemnts.Sum(b => b.Amount);
                    worksheet.Cells[rowNumber, 5].Value = item.Amount - item.Assignemnts.Sum(b => b.Amount);
                    worksheet.Cells[rowNumber, 6].Value = item.ReceiveBillType.GetDisplayName();
                    worksheet.Cells[rowNumber, 7].Value = item.CurrencyType.GetDisplayName();
                    worksheet.Cells[rowNumber, 8].Value = new PersianDateTime(item.CreationTime).ToShortDateString();
                    rowNumber++;
                }
                //worksheet.Column(1).Width = 5;
                //worksheet.Column(2).Width += 10;
                //worksheet.Column(3).Width += 10;
                //worksheet.Column(4).Width += 70;
                //worksheet.Column(5).Width += 30;
                //worksheet.Column(6).Width += 10;
                //worksheet.Column(7).Width += 10;
                //worksheet.Column(8).Width += 10;
                //worksheet.Column(9).Width += 10;
                //worksheet.Column(10).Width += 10;

                package.Workbook.Properties.Title = "Portal";
                package.Workbook.Properties.Author = "Dapna.Co";
                package.Workbook.Properties.Company = "MSVCO";

                var stream = new MemoryStream();
                package.SaveAs(stream);
                stream.Position = 0;

                const string contentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                stream.Position = 0;
                return File(stream, contentType, fileName);
            }
        }



        [DontWrapResult]
        public async Task<ActionResult> ReceiveBillTransactionRead([DataSourceRequest] DataSourceRequest request, int id)
        {
            var items = await _transactionAppService.GetReceiveBillTransactions(id);
            var paymentrequests = items.Select(a => new TransactionViewModel()
            {
                Id = a.Id,
                Amount = a.Amount,
                CurrencyType = a.CurrencyType,
                Description = a.Description,
                CreationTime = new PersianDateTime(a.CreationTime).ToShortDateString(),
                ProjectName = a.Project != null ? a.Project.Title : "-",
                TransactionType = a.TransactionType
            });
            var dsResult = paymentrequests.ToDataSourceResult(request);
            return Json(dsResult, new JsonSerializerSettings() { ContractResolver = new DefaultContractResolver() });
        }


        [AcceptVerbs("Post")]
        [DontWrapResult]
        [DisableValidation]
        public async Task<IActionResult> ReceiveBillTransactionDestroy([DataSourceRequest] DataSourceRequest request, TransactionViewModel item)
        {
            if (item != null)
            {
                await _transactionAppService.Delete(new EntityDto(item.Id));
            }
            return Json(new[] { item }.ToDataSourceResult(request, ModelState));
        }



        [AcceptVerbs("Post")]
        [DontWrapResult]
        [DisableValidation]
        public async Task<IActionResult> ReceiveBillTransactionUpdate([DataSourceRequest] DataSourceRequest request, TransactionViewModel item)
        {
            if (item != null)
            {
                var dto = await _transactionAppService.Get(new EntityDto<int>(item.Id));
                dto.Amount = item.Amount;
                await _transactionAppService.Update(dto);
            }
            return Json(new[] { item }.ToDataSourceResult(request, ModelState));
        }


        [HttpPost]
        [DontWrapResult]
        [DisableValidation]
        public async Task<IActionResult> AddReceiveBillTransaction(string amount, string description, int recevieBillId, int projectId, string date)
        {
            try
            {
                var recevieBill = await _receiveBillAppService.Get(new EntityDto<int>(recevieBillId));
                var item = new TransactionDto()
                {
                    TransactionType = Enums.TransactionType.Receive,
                    Amount = Convert.ToDecimal(amount),
                    CurrencyType = recevieBill.CurrencyType,
                    PaymentType = Enums.PaymentType.Cash,
                    Description = description,
                    ReceiveBillId = recevieBill.Id,
                    ProjectId = projectId,
                    CreationTime = PersianDateTime.Parse(date)
                };

                await _transactionAppService.Create(item);

                return Json(new { status = "success" });
            }
            catch (Exception ex)
            {
                return Json(new { status = "error", message = ex.Message });
                throw;
            }
        }



        [AcceptVerbs("Post")]
        [DontWrapResult]
        [DisableValidation]
        public async Task<IActionResult> ReceiveBillsDestroy([DataSourceRequest] DataSourceRequest request, ReceiveBillViewModel item)
        {
            var result = await _receiveBillAppService.DeleteReceiveBill(item.Id);
            if (!result)
            {
                ModelState.AddModelError("DependentRecord", "-------");
            }
            return Json(new[] { item }.ToDataSourceResult(request, ModelState));
        }



        [HttpPost]
        [DontWrapResult]
        [DisableValidation]
        public async Task<IActionResult> EditTransaction(int modal_id, string modal_amount, string modal_date)
        {
            try
            {
                var transactionDto = await _transactionAppService.Get(new EntityDto<int>(modal_id));
                transactionDto.Amount = Convert.ToDecimal(modal_amount);
                transactionDto.CreationTime = PersianDateTime.Parse(modal_date);
                await _transactionAppService.Update(transactionDto);

                return Json(new { status = "success" });
            }
            catch (Exception ex)
            {
                return Json(new { status = "error", message = ex.Message });
                throw;
            }
        }


        public IActionResult Contractors()
        {
            return View();
        }


        [DontWrapResult]
        public async Task<ActionResult> ContractorRead([DataSourceRequest] DataSourceRequest request)
        {
            var items = await _contractorAppService.GetAllContractors();

            var itemResults = items.Select(a => new ContractorViewModel()
            {
                Id = a.Id,
                Code = a.Code,
                Name = a.Name
            }).ToList();

            var dsResult = itemResults.ToDataSourceResult(request);
            return Json(dsResult, new JsonSerializerSettings() { ContractResolver = new DefaultContractResolver() });
        }


        [HttpPost]
        [DontWrapResult]
        [DisableValidation]
        public async Task<IActionResult> AddContractor(int? Id, string Name, string Code)
        {
            try
            {
                if (Id != null)
                {
                    await _contractorAppService.Update(new ContractorDto()
                    {
                        Id = Id.Value,
                        Name = Name,
                        Code = Code
                    });
                }
                else
                {
                    await _contractorAppService.Create(new ContractorDto()
                    {
                        Name = Name,
                        Code = Code
                    });
                }


                return Json(new { status = "success" });
            }
            catch (Exception ex)
            {
                return Json(new { status = "error", message = ex.Message });
                throw;
            }
        }


        [AcceptVerbs("Post")]
        [DontWrapResult]
        [DisableValidation]
        public async Task<IActionResult> ContractorDestroy([DataSourceRequest] DataSourceRequest request, ContractorViewModel item)
        {
            await _contractorAppService.Delete(new EntityDto<int>(item.Id));
            return Json(new[] { item }.ToDataSourceResult(request, ModelState));
        }


        [HttpPost]
        [DontWrapResult]
        [DisableValidation]
        public async Task<IActionResult> AddProjectToUser(long userId, int projectId)
        {
            try
            {
                var isDuplicate = await _project_User_MappingAppService.IsDuplicate(userId, projectId);
                if (!isDuplicate)
                {
                    await _project_User_MappingAppService.Create(new Project_User_MappingDto()
                    {
                        ProjectId = projectId,
                        UserId = userId
                    });
                    return Json(new { status = "success" });
                }
                else
                {
                    return Json(new { status = "duplicate" });
                }

            }
            catch (Exception ex)
            {
                return Json(new { status = "error", message = ex.Message });
                throw;
            }
        }


        [DontWrapResult]
        public async Task<ActionResult> UserProjectRead([DataSourceRequest] DataSourceRequest request, long id)
        {
            var items = await _project_User_MappingAppService.GetUserProjects(id);
            var paymentrequests = items.Select(a => new UserProjectViewModel()
            {
                Id = a.Id,
                ProjectName = a.Project.Title
            });
            var dsResult = paymentrequests.ToDataSourceResult(request);
            return Json(dsResult, new JsonSerializerSettings() { ContractResolver = new DefaultContractResolver() });
        }


        [AcceptVerbs("Post")]
        [DontWrapResult]
        [DisableValidation]
        public async Task<IActionResult> UserProjectDestroy([DataSourceRequest] DataSourceRequest request, UserProjectViewModel item)
        {
            if (item != null)
            {
                await _project_User_MappingAppService.Delete(new EntityDto(item.Id));
            }
            return Json(new[] { item }.ToDataSourceResult(request, ModelState));
        }




        [HttpPost]
        [DontWrapResult]
        [DisableValidation]
        public async Task<IActionResult> PaymentRequestChangeStatus(int paymentRequestId, int status)
        {
            try
            {
                var paymentRequest = await _paymentRequestAppService.GetWillDetails(paymentRequestId);
                paymentRequest.PaymentRequestStatus = (PaymentRequestStatus)status;
                if (status == 2)
                {
                    var urlWithAccessToken = "https://hooks.slack.com/services/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";
                    var client = new SlackClient(urlWithAccessToken);
                    var message = String.Empty;
                    message = String.Format("ثبت درخواست پرداخت جدید" + "\n" +
                                               "ثبت کننده : " + _userManager.GetUserByIdAsync(paymentRequest.CreatorUserId.Value).Result.FullName + "\n" +
                                               "پروژه : " + paymentRequest.Project.Title + "\n" +
                                               "مبلغ درخواستی : " + paymentRequest.Amount + "\n" +
                                               "در وجه : " + paymentRequest.PayTo + "\n" +
                                               "شرح : " + paymentRequest.Description + "\n"
                                               );
                    client.Send(message);
                }
                var itemDto = ObjectMapper.Map<PaymentRequestDto>(paymentRequest);
                await _paymentRequestAppService.Update(itemDto);

                return Json(new { status = "success" });
            }
            catch (Exception ex)
            {
                return Json(new { status = "error", message = ex.Message });
                throw;
            }
        }


        [DontWrapResult]
        public async Task<ActionResult> Katibe_DarkhastRead([DataSourceRequest] DataSourceRequest request)
        {
            try
            {
                var userId = AbpSession.UserId.Value;
                var userProjects = await _project_User_MappingAppService.GetUserProjects(userId);

                var codes = userProjects.Where(a => a.Project != null && !string.IsNullOrEmpty(a.Project.Code)).Select(a => a.Project.Code).ToList();

                var items = _katibe_DarkhastAppService.GetAllKatibe_Darkhasts(codes);

                var itemResults = items.Select(a => new KatibeDarkhastViewModel()
                {
                    Id = a.Id,
                    DarVajh = a.DarVajh,
                    Kharid = a.Kharid,
                    Onvan = a.Onvan,
                    CodeOnvan = a.CodeOnvan,
                    ShenaseOnvan = a.ShenaseOnvan,
                    TaaedMali = a.TaaedMali,
                    Shenasname = a.Shenasname != null ? a.Shenasname.ToString() : "0"
                }).ToList();

                var dsResult = itemResults.ToDataSourceResult(request);
                return Json(dsResult, new JsonSerializerSettings() { ContractResolver = new DefaultContractResolver() });
            }
            catch (Exception ex)
            {
                var message = ex.Message;
                throw;
            }

        }

        public async Task<IActionResult> RemainCredit()
        {
            var Items = await _project_User_MappingAppService.GetUserProjects(AbpSession.UserId.Value);
            ViewBag.ProjectID = Items.Select(a => new ProjectViewModel()
            {
                Id = a.ProjectId,
                Title = a.Project.Title
            }).ToList();

            return View();
        }

        [DontWrapResult]
        public async Task<IActionResult> RemainCreditTableData([DataSourceRequest] DataSourceRequest request)
        {
            var List = new List<RemainCreditViewModel>();

            var UsersProjects = _project_User_MappingAppService.GetUserProjects(AbpSession.UserId.Value).Result.Select(a => a.ProjectId);

                var Result = await _remainCreditAppService.GetUsersAllProjects(UsersProjects.ToList());

                List.AddRange(Result.Select(Record => new RemainCreditViewModel
                {
                    ProjectID = Record.ProjectID,
                    ProjectCode = Record.ProjectCode,
                    ProjectName = Record.ProjectName,
                    FourthLevelCode = Record.FourthLevelCode,
                    PayTo = Record.PayTo,
                    RemainDebtAmount = String.Format("{0:n0}", Record.RemainDebtAmount),
                    RemainCreditAmount = String.Format("{0:n0}", Record.RemainCreditAmount) ,
                    UploadDate = Record.CreationTime != null ? new PersianDateTime(Record.CreationTime).ToShortDateString() : ""
                }));

            var dsResult = List.ToDataSourceResult(request);
            return Json(dsResult, new JsonSerializerSettings() { ContractResolver = new DefaultContractResolver() });
        }


        [AcceptVerbs("Post")]
        [DontWrapResult]
        [DisableValidation]
        public IActionResult RemainCreditFilter(int[] ProjectIDs, string PayTo, string FourthLevelCode, int SortType , int ShowAll) //ShowAll Is Flag That Represent That If User Wants To See The Values That Has 0 CreditAmount And DebtAmount
        {
            var UsersProjects = _project_User_MappingAppService.GetUserProjects(AbpSession.UserId.Value).Result.Select(a => a.ProjectId);

            var Result = _remainCreditAppService.GetUsersFilteredProjects(UsersProjects.ToList(), ProjectIDs.ToList(), PayTo, FourthLevelCode , SortType, ShowAll).Result;

            var List = new List<RemainCreditViewModel>();

            foreach (var Item in Result)
            {
                List.Add(new RemainCreditViewModel
                {
                    ProjectID = Item.ProjectID,
                    ProjectCode = Item.ProjectCode,
                    ProjectName = Item.ProjectName,
                    FourthLevelCode = Item.FourthLevelCode,
                    PayTo = Item.PayTo,
                    RemainDebtAmount = String.Format("{0:n0}", Item.RemainDebtAmount),
                    RemainCreditAmount = String.Format("{0:n0}", Item.RemainCreditAmount),
                    UploadDate = Item.CreationTime != null ? new PersianDateTime(Item.CreationTime).ToShortDateString() : ""
                });
            }

            return Json(List, new JsonSerializerSettings() { ContractResolver = new DefaultContractResolver() });
        }

        public async Task<IActionResult> TotalCreditAndDebt(int?[] ProjectIDs , string PayTo , string FourthLevelCode)
        {
            long TotalDebt = 0;
            long TotalCredit = 0;
            var List = new List<RemainCreditViewModel>();

            var UsersProjects = _project_User_MappingAppService.GetUserProjects(AbpSession.UserId.Value).Result.Select(a => a.ProjectId);

            if (ProjectIDs.Count() > 0  || PayTo != null || FourthLevelCode != null)
            {
                var Result = await _remainCreditAppService.GetFilteredTotalDebtAndCredit(UsersProjects.ToList() , ProjectIDs.ToList() , PayTo, FourthLevelCode);

                foreach (var item in Result)
                {
                    TotalDebt = (long)(TotalDebt + item.RemainDebtAmount);
                    TotalCredit = (long)(TotalCredit + item.RemainCreditAmount);
                }

                List.Add(new RemainCreditViewModel { TotalDebtAmount = String.Format("{0:n0}", TotalDebt), TotalCreditAmount = String.Format("{0:n0}", TotalCredit) });
            }

            else
            {
                    var Result = await _remainCreditAppService.GetTotalDebtAndCredit(UsersProjects.ToList());

                    foreach (var item in Result)
                    {
                        TotalDebt = (long)(TotalDebt + item.RemainDebtAmount);
                        TotalCredit = (long)(TotalCredit + item.RemainCreditAmount);
                    }
                List.Add(new RemainCreditViewModel { TotalDebtAmount = String.Format("{0:n0}", TotalDebt), TotalCreditAmount = String.Format("{0:n0}", TotalCredit) });
            }

            return Json(List);
        }

        public async Task<IActionResult> RemainCreditImportExcel(IFormFile ExcelFile)
        {
            string FilePath = "";

            var Model = new RemainCreditViewModel();

            SqlConnection SQLConnection = new SqlConnection("Server=192.168.2.47; Database=msv_portal; user id=msvportal_user; password=Msv8810;Trusted_Connection=false");

            SQLConnection.Open();

            if (!System.IO.Directory.Exists(_hostingEnvironment.WebRootPath + "\\Excel")) { System.IO.Directory.CreateDirectory(_hostingEnvironment.WebRootPath + "\\Excel"); } //Checks If Excel's Upload Folder Exists - If It Doesn't Then Create It

            if (!System.IO.Directory.Exists(_hostingEnvironment.WebRootPath + "\\Excel" + "\\RemainCredit-ImportExcel")) { System.IO.Directory.CreateDirectory(_hostingEnvironment.WebRootPath + "\\Excel" + "\\RemainCredit-ImportExcel"); } //Checks If Excel's Upload Folder Exists - If It Doesn't Then Create It

            if (ExcelFile != null)
            {
                //Save The ExcelFile In Excel Folder In WWWRoot Path

                string Excels = Path.Combine(_hostingEnvironment.WebRootPath, "Excel");
                string RemainCreditExcels = Path.Combine(Excels, "RemainCredit-ImportExcel");

                if (ExcelFile.FileName.Contains(".xlsx")) { FilePath = Path.Combine(RemainCreditExcels, ExcelFile.FileName.Replace(".xlsx", " - " + new PersianDateTime(DateTime.Parse(DateTime.Now.ToString())).ToLongDateTimeInt() + ".xlsx")); } //Add DateTime To Imported Excel File
                else if (ExcelFile.FileName.Contains(".xls")) { FilePath = Path.Combine(RemainCreditExcels, ExcelFile.FileName.Replace(".xls", " - " + new PersianDateTime(DateTime.Parse(DateTime.Now.ToString())).ToLongDateTimeInt() + ".xls")); } //Add DateTime To Imported Excel File
                else { return Json("فایل مورد نظر از نوع اکسل نمی باشد"); } // If File Wasn't Excel Format , Returns Error Alert

                using (Stream FileStream = new FileStream(FilePath, FileMode.Create)) { await ExcelFile.CopyToAsync(FileStream); } //Copy Imported Excel File In Excel Folder In WWWRoot 
            }


            FileStream Stream = System.IO.File.Open(FilePath, FileMode.Open, FileAccess.Read);
            IExcelDataReader ExcelReader;


            //1.Reading Excel file
            if (Path.GetExtension(FilePath).ToUpper() == ".XLS")
            {
                //1.1 Reading from a binary Excel file('97-2003 format; *.xls)
                ExcelReader = ExcelReaderFactory.CreateBinaryReader(Stream);
            }
            else
            {
                //1.2 Reading from a OpenXml Excel file(2007 format; *.xlsx)
                ExcelReader = ExcelReaderFactory.CreateOpenXmlReader(Stream);
            }

            //2.DataSet - The result of each spreadsheet will be created in the result.Tables
            DataSet Result = ExcelReader.AsDataSet();
            DataTable DataTable = Result.Tables[0];

            var CountRows = DataTable.Rows.Count; //Number Of All The Rows In Excel 

            //Delete Whole RemainCredit Table Then Start To Import New Data To RemainCredit Table
            var DeleteQuery = "DELETE FROM RemainCredits";
            SqlCommand DeleteCommand = new SqlCommand(DeleteQuery, SQLConnection);
            DeleteCommand.ExecuteNonQuery();

            for (int Index = 2; Index < CountRows; Index++)
            {
                var ProjectID = _projectAppService.GetProjectByProjectCode(Convert.ToString(DataTable.Rows[Index][0])).Result.Select(a => a.Id).FirstOrDefault();
                var ProjectCode = Convert.ToString(DataTable.Rows[Index][0]);
                var ProjectName = _projectAppService.GetProjectByProjectCode(Convert.ToString(DataTable.Rows[Index][0])).Result.Select(a => a.Title).FirstOrDefault();
                var FourthLevelCode = Convert.ToString(DataTable.Rows[Index][2]);
                var PayTo = Convert.ToString(DataTable.Rows[Index][3]);
                var RemainDebtAmount = Convert.ToInt64(DataTable.Rows[Index][4]);
                var RemainCreditAmount = Convert.ToInt64(DataTable.Rows[Index][5]);

                var Query = "INSERT INTO RemainCredits (CreationTime, IsDeleted , ProjectID , ProjectCode , ProjectName , FourthLevelCode , PayTo , RemainCreditAmount , RemainDebtAmount) VALUES ('" + DateTime.Parse(DateTime.Now.ToString()) + "' , " + 0 + " , " + ProjectID +" , '" + ProjectCode + "' , N'" + ProjectName + "' , '" + FourthLevelCode + "' , N'" + PayTo + "' , " + RemainCreditAmount + " , " + RemainDebtAmount + ")";
                SqlCommand Command = new SqlCommand(Query, SQLConnection);
                Command.ExecuteNonQuery();
            }

            SQLConnection.Close();

            Stream.Close();

            return Json("اکسل با موفقیت وارد شد");
        }

        public async Task<ActionResult> RemindCreditExportExcel(string ProjectIDs , string PayTo , string FourthLevelCode , int SortType , int ShowAll) //ShowAll Is Flag That Represent That If User Wants To See The Values That Has 0 CreditAmount And DebtAmount
        {
            var Result = new List<RemainCredit>();
            var List = new List<RemainCredit>();

            int[] ProjectIDsList = ProjectIDs != null ? ProjectIDs.Split(",").Select(int.Parse).ToArray() : new int[] { };

            var UsersProjects = _project_User_MappingAppService.GetUserProjects(AbpSession.UserId.Value).Result.Select(a => a.ProjectId);

            if (ProjectIDs != null || PayTo != null || FourthLevelCode != null || SortType != 0 || ShowAll == 1 )
            {
                List = await _remainCreditAppService.GetUsersFilteredProjects(UsersProjects.ToList() , ProjectIDsList.ToList(), PayTo, FourthLevelCode , SortType , ShowAll);
            }

            else
            {
                    Result = await _remainCreditAppService.GetUsersAllProjects(UsersProjects.ToList());

                    List.AddRange(Result.Select(Record => new RemainCredit
                    {
                        ProjectID = Record.ProjectID,
                        ProjectCode = Record.ProjectCode,
                        ProjectName = Record.ProjectName,
                        FourthLevelCode = Record.FourthLevelCode,
                        PayTo = Record.PayTo,
                        RemainDebtAmount = Record.RemainDebtAmount,
                        RemainCreditAmount = Record.RemainCreditAmount,
                        //UploadDate = Record.CreationTime != null ? new PersianDateTime(Record.CreationTime).ToShortDateString() : ""
                    }));
            }

            if (!System.IO.Directory.Exists(_hostingEnvironment.WebRootPath + "\\Excel")) { System.IO.Directory.CreateDirectory(_hostingEnvironment.WebRootPath + "\\Excel"); } //Checks If Excel's Upload Folder Exists - If It Doesn't Then Create It

            if (!System.IO.Directory.Exists(_hostingEnvironment.WebRootPath + "\\Excel\\RemainCredit-ExportExcel")) { System.IO.Directory.CreateDirectory(_hostingEnvironment.WebRootPath + "\\Excel\\RemainCredit-ExportExcel"); } //Checks If Excel's Upload Folder Exists - If It Doesn't Then Create It

            // Create the file using the FileInfo object
            var FileName = "مانده بستانکار - " + new PersianDateTime(DateTime.Parse(DateTime.Now.ToString())).ToLongDateTimeInt() + ".xlsx";
            var FileNameForReturn = "مانده بستانکار - " + new PersianDateTime(DateTime.Parse(DateTime.Now.ToString())).ToShortDateString() + ".xlsx"; //This Format Of File Is Used At The End For Returning The File 
            var OutPutDirectory = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot/Excel/RemainCredit-ExportExcel/");

            var ExcelFile =  new FileInfo(OutPutDirectory + FileName);

            using (var Package = new ExcelPackage(ExcelFile))
            {
                //Save The ExcelFile In Excel Folder In WWWRoot Path=
                string Excels = Path.Combine(_hostingEnvironment.WebRootPath, "Excel");
                string FilePath = Path.Combine(Excels, "RemainCredit-ExportExcel");

                //ExcelFile.Create(); 
                //ExcelFile.CopyTo(FilePath);  //Copy ExportExcel File In Excel Folder In WWWRoot 

                // add a new worksheet to the empty workbook
                ExcelWorksheet worksheet = Package.Workbook.Worksheets.Add("RemainCredit - " + DateTime.Now.ToShortDateString());
                worksheet.HeaderFooter.FirstFooter.LeftAlignedText = string.Format("Generated: {0}", DateTime.Now.ToShortDateString());
                worksheet.Row(1).Height = 15;
                worksheet.View.RightToLeft = true;

                // Start adding the header
                worksheet.Cells[1, 1].Value = "کد پروژه";
                worksheet.Cells[1, 2].Value = "نام پروژه";
                worksheet.Cells[1, 3].Value = "تفصیل";
                worksheet.Cells[1, 4].Value = "نام شخص";
                worksheet.Cells[1, 5].Value = "مانده بدهکار";
                worksheet.Cells[1, 6].Value = "مانده بستانکار";


                //// Add the second row of header data
                using (var range = worksheet.Cells[1, 1, 1, 6])
                {
                    range.Style.Font.Bold = true;
                    range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    range.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.DarkBlue);
                    range.Style.Font.Color.SetColor(System.Drawing.Color.WhiteSmoke);
                    range.Style.ShrinkToFit = false;
                    range.Style.WrapText = true;
                    range.Style.Locked = true;
                }

                int rowNumber = 2;

                foreach (var Item in List)
                {
                    worksheet.Cells[rowNumber, 1].Value = Item.ProjectCode;
                    worksheet.Cells[rowNumber, 2].Value = Item.ProjectName;
                    worksheet.Cells[rowNumber, 3].Value = Item.FourthLevelCode;
                    worksheet.Cells[rowNumber, 4].Value = Item.PayTo;
                    worksheet.Cells[rowNumber, 5].Value = Item.RemainDebtAmount;
                    worksheet.Cells[rowNumber, 6].Value = Item.RemainCreditAmount;

                    rowNumber++;
                }

                Package.Workbook.Properties.Title = "RemainCredit - ExportExcel";
                Package.Workbook.Properties.Author = "Dapna.CO";
                Package.Workbook.Properties.Company = "MSVCO";

                var stream = new MemoryStream();
                Package.SaveAs(stream);
                stream.Position = 0;

                const string contentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                stream.Position = 0;
                return File(stream, contentType, FileNameForReturn);
            }
        }
    }
}