using GJWalker.DBAccess;
using GJWalker.Model;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Identity;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Rendering;
using Microsoft.Extensions.Configuration;
using Newtonsoft.Json;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.Linq;
using System.Security.Claims;
using System.Text;
using System.Threading.Tasks;

namespace GJWalker.Controllers
{
    [Authorize]
    public class MainController : Controller
    {
        modGJWFCU.GJWFCU objGJWFCU = new modGJWFCU.GJWFCU();
        private IConfiguration Configuration;
        private string conn = "";
        private string connCP4 = "";
        private int UserID = 0;
        private readonly UserManager<IdentityUser> _userManager;
        private readonly IHttpContextAccessor _httpContextAccessor;
        //private modGJWFCU.GJWFCU gjwFCU = new modGJWFCU.GJWFCU();
        public MainController(IConfiguration _configuration, UserManager<IdentityUser> userManager, IHttpContextAccessor httpContextAccessor)
        {
            _userManager = userManager;
            _httpContextAccessor = httpContextAccessor;
            Configuration = _configuration;
            conn = Configuration.GetConnectionString("DefaultConnection");
            connCP4 = Configuration.GetConnectionString("CP4Connection");
            var userId = _httpContextAccessor.HttpContext.User.FindFirst(ClaimTypes.NameIdentifier).Value;
            UserID =DbAccess.GetUserID(conn,userId);
        }
        public IActionResult Index()
        {
            return View();
        }
        public ActionResult GetProductById(int id)
        {
            var products = DbAccess.GetSampleProducts().Where(x => x.id == id); ;

            if (products != null)
            {
                Product model = new Product();

                foreach (var item in products)
                {
                    model.name = item.name;
                    model.price = item.price;
                    model.department = item.department;
                }

                return PartialView("_GridEditPartial", model);
            }

            return View();
        }
        public ActionResult GetProducts(string sidx, string sord, int page, int rows)
        {
            List<Product> products = DbAccess.GetSampleProducts();
            int pageIndex = Convert.ToInt32(page) - 1;
            int pageSize = rows;
            int totalRecords = products.Count();
            int totalPages = (int)Math.Ceiling((float)totalRecords / (float)pageSize);

            var data = products.OrderBy(x => x.id)
                          .Skip(pageSize * (page - 1))
                          .Take(pageSize).ToList();

            var jsonData = new
            {
                total = totalPages,
                page = page,
                records = totalRecords,
                rows = data
            };

            return Json(jsonData);
        }
        public IActionResult Selection()
        {
            List<ProjectDDLModel> cl = new List<ProjectDDLModel>();
            cl = DbAccess.GetProjectDDL(conn, UserID.ToString());
            cl.Insert(0, new ProjectDDLModel { ProjectID = 0, Project = "--Select Project Name--" });
            ViewBag.Projects = new SelectList(cl, "ProjectID", "Project"); ;

            return View();
        }
        [HttpPost]
        public ActionResult GetUnitByProjectId(string projectId)
        {
            ProjectModel project = new ProjectModel();
            project = DbAccess.GetProjectById(conn, projectId);
            List<UnitModel> objUnit = new List<UnitModel>();
            objUnit = DbAccess.GetUnitDDL(conn, projectId);
            objUnit.Insert(0, new UnitModel { UnitID = 0, Tag = "--Select Unit Name--" });
            SelectList unit = new SelectList(objUnit, "UnitID", "Tag");
            ProjectUnitModel projectUnit = new ProjectUnitModel();
            projectUnit.Project = project;
            projectUnit.Unit = unit;
            return Json(projectUnit);
        }
        [HttpPost]
        public ActionResult GetUnitById(string unitId)
        {
            //objGJWFCU.ClearOptions(conn, Convert.ToInt32(unitId));
            UnitResult dt = DbAccess.GetUnitByID(Configuration.GetConnectionString("DefaultConnection"), unitId);
            return Json(dt);
        }

        [HttpPost]
        public ActionResult GetPrintUnitById(string unitId)
        {
            DataTable dt = DbAccess.GetPrintUnitById(Configuration.GetConnectionString("DefaultConnection"), unitId);
            List<UnitPrintResult> options = DbAccess.CreateListFromTable<UnitPrintResult>(dt);
            return Json(dt);
        }

        [HttpPost]
        public ActionResult GetOptionData(string unitId)
        {
            DataTable dtOption = DbAccess.GetOption(conn, unitId);
            List<OptionModel> options = DbAccess.CreateListFromTable<OptionModel>(dtOption);
            return PartialView("_optionGrid", options);
        }

        [HttpPost]
        public ActionResult GetCoilData(string modelId, string unitId)
        {
            DataTable dtOption = DbAccess.GetCoilData(conn, modelId, unitId);
            List<ColiDataModel> options = DbAccess.CreateListFromTable<ColiDataModel>(dtOption);
            return PartialView("_coilData", options);
        }
        [HttpPost]
        public ActionResult GetCoilSelectionData(string modelId, string unitId)
        {
            DataTable dtOption = DbAccess.GetCoilSelectionData(conn, modelId, unitId);
            List<ColiDataModel> options = DbAccess.CreateListFromTable<ColiDataModel>(dtOption);
            return PartialView("_coilData", options);
        }
        [HttpPost]
        public ActionResult GetFanData(string modelId, string unitId)
        {
            DataTable dtOption = DbAccess.GetFanData(conn, modelId, unitId);
            List<FanDataModel> options = DbAccess.CreateListFromTable<FanDataModel>(dtOption);
            if (options != null)
            {
                if (options[0] != null)
                {
                    DataTable curveData = objGJWFCU.GetFanSysCurve(conn, Convert.ToInt64(unitId), Convert.ToInt64(options[0].ModelID));
                    List<DataPoint> fanPress = new List<DataPoint>();
                    List<DataPoint> sysPress = new List<DataPoint>();
                    List<FanCurveData> fanCurveData = DbAccess.CreateListFromTable<FanCurveData>(curveData).OrderBy(p=>p.Airflow).Distinct().ToList();
                    int cnt = 100;
                    foreach (var item in fanCurveData)
                    {
                        fanPress.Add(new DataPoint(item.Airflow, Convert.ToDouble(item.FanPress), "FanPress"));
                        sysPress.Add(new DataPoint(item.Airflow, Convert.ToDouble(item.SysPress), "SysPress"));
                        cnt=cnt+100;
                    }
                    options[0].FanPress = JsonConvert.SerializeObject(fanPress);
                    options[0].SysPress = JsonConvert.SerializeObject(sysPress);
                }
            }

            return PartialView("_fanData", options[0]);
        }

        [HttpPost]
        public ActionResult GetFanSelectionData(string modelId, string unitId)
        {
            DataTable dtOption = DbAccess.GetFanSelectionData(conn, modelId, unitId);
            List<FanDataModel> options = DbAccess.CreateListFromTable<FanDataModel>(dtOption);
            if (options != null)
            {
                if (options[0] != null)
                {
                    DataTable curveData = objGJWFCU.GetFanSysCurve(conn, Convert.ToInt64(unitId), Convert.ToInt64(options[0].ModelID));
                    List<DataPoint> fanPress = new List<DataPoint>();
                    List<DataPoint> sysPress = new List<DataPoint>();
                    List<FanCurveData> fanCurveData = DbAccess.CreateListFromTable<FanCurveData>(curveData).OrderBy(p => p.Airflow).Distinct().ToList();
                    int cnt = 100;
                    foreach (var item in fanCurveData)
                    {
                        fanPress.Add(new DataPoint(item.Airflow, Convert.ToDouble(item.FanPress), "FanPress"));
                        sysPress.Add(new DataPoint(item.Airflow, Convert.ToDouble(item.SysPress), "SysPress"));
                        cnt = cnt + 100;
                    }
                    options[0].FanPress = JsonConvert.SerializeObject(fanPress);
                    options[0].SysPress = JsonConvert.SerializeObject(sysPress);
                }
            }

            return PartialView("_fanData", options[0]);
        }

        [HttpPost]
        public ActionResult GetAttributeData(string modelId, string unitId)
        {
            DataTable dtOption = DbAccess.GetAttributeData(conn, modelId, unitId);
            List<AttributeModel> options = DbAccess.CreateListFromTable<AttributeModel>(dtOption);
            return PartialView("_attribute", options[0]);
        }

        [HttpPost]
        public ActionResult GetAttributeSelectionData(string modelId, string unitId)
        {
            DataTable dtOption = DbAccess.GetAttributeSelectionData(conn, modelId, unitId);
            List<AttributeModel> options = DbAccess.CreateListFromTable<AttributeModel>(dtOption);
            return PartialView("_attribute", options[0]);
        }


        [HttpPost]
        public ActionResult GetAncillariesData(string modelId, string unitId)
        {
            DataTable dtOption = DbAccess.GetAncillariesData(conn, modelId, unitId);
            List<AncillariesModel> options = DbAccess.CreateListFromTable<AncillariesModel>(dtOption);
            return PartialView("_ancillaries", options[0]);
        }

        [HttpPost]
        public ActionResult GetAncillariesSelectionData(string modelId, string unitId)
        {
            DataTable dtOption = DbAccess.GetAncillariesSelectionData(conn, modelId, unitId);
            List<AncillariesModel> options = DbAccess.CreateListFromTable<AncillariesModel>(dtOption);
            return PartialView("_ancillaries", options[0]);
        }

        [HttpPost]
        public ActionResult GetModelImages(string modelId, string unitId)
        {
            DataTable dtOption = DbAccess.GetModelImages(conn, modelId, unitId);
            List<ModelImages> options = DbAccess.CreateListFromTable<ModelImages>(dtOption);
            return Json(options[0]);
        }

        [HttpPost]
        public ActionResult GetSelectionModelImages(string modelId, string unitId)
        {
            DataTable dtOption = DbAccess.GetSelectionModelImages(conn, modelId, unitId);
            List<ModelImages> options = DbAccess.CreateListFromTable<ModelImages>(dtOption);
            return Json(options[0]);
        }

        [HttpPost]
        public ActionResult SelectFCU(UnitResult obj)
        {
            OptionDataRequest optionsRequest = new OptionDataRequest();
            try
            {
                var isSaved = DbAccess.UpdateUnit(conn, obj);
                //Update input  data in table 
                BitArray myErrors = objGJWFCU.GetInputErrors(conn, obj.ProjectID, obj.UnitID);
                string errMsg = "";
                for (int i = 0; i < myErrors.Count; i++)
                {
                    if (myErrors[i])
                    {
                        var msg = objGJWFCU.GetErrorMessage(conn, obj.ProjectID, i);
                        if (errMsg != "")
                            errMsg = errMsg + "," + msg;
                        else
                            errMsg = msg;
                    }
                }

                if (errMsg == "")
                {
                    int selCount = objGJWFCU.PopulateOptions(conn,connCP4, obj.ProjectID, obj.UnitID);
                    if (selCount != 0)
                    {
                        optionsRequest.Status = "Success";
                        optionsRequest.ErrorMessage = "";
                    }
                    else if (selCount == 0)
                    {

                    }
                    else
                    {

                    }
                }
                else
                {

                    optionsRequest.Status = "Failure";
                    optionsRequest.ErrorMessage = errMsg;
                }
            }
            catch (NullReferenceException ex)
            {

                optionsRequest.Status = "Failure";
                optionsRequest.ErrorMessage = ex.Message;
            }
            catch (Exception ex)
            {

                optionsRequest.Status = "Failure";
                optionsRequest.ErrorMessage = ex.Message;
            }

            return Json(optionsRequest);
        }
        public IActionResult Schedule(string u, string p)
        {
            string? projectName = "";

            string? Date = null;
            string? OrganisedBy = "";
            string? SelectedBy = "";
            string? Email = "";
            string? QuoteNumber = "";
            string? Revision = "";
            string? AcctManager = "";
            Single? Discount = 0;
            string? ShipTo = "";
            string? Notes = "";
            List<ProjectDDLModel> cl = new List<ProjectDDLModel>();
            cl = DbAccess.GetProjectDDL(conn, UserID.ToString());
            cl.Insert(0, new ProjectDDLModel { ProjectID = 0, Project = "--Select Project Name--" });
            ViewBag.Projects = new SelectList(cl, "ProjectID", "Project"); ;
            if (!String.IsNullOrEmpty(p))
            {
                ProjectModel project= DbAccess.GetProjectById(conn, p);
                if(project!=null)
                {
                    projectName = project.ProjectName;
                    Date = project.Date;
                    OrganisedBy=project.OrganisedBy;
                    SelectedBy=project.SelectedBy;
                    Email=project.Email;
                    QuoteNumber=project.QuoteNumber;
                    Revision=project.Revision;
                    AcctManager=project.AcctManager;
                    Discount=project.Discount;
                    ShipTo=project.ShipTo;
                    Notes=project.Notes;
                }
            }
            ViewBag.ProjectName=projectName;
          
            ViewBag.Date =(Date==""?DateTime.Now.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture): Date);
            ViewBag.OrganisedBy = OrganisedBy;
            ViewBag.SelectedBy = SelectedBy;
            ViewBag.Email = Email;
            ViewBag.QuoteNumber = QuoteNumber;
            ViewBag.Revision = Revision;
            ViewBag.AcctManager = AcctManager;
            ViewBag.Discount = Discount;
            ViewBag.ShipTo = ShipTo;
            ViewBag.Notes = Notes;

            DataTable dtOption = DbAccess.GetSelectionData(conn,p);
            List<SelectionModel> options = DbAccess.CreateListFromTable<SelectionModel>(dtOption);
            return View(options);
        }
        public ActionResult GetPdf(string filename)
        {
            return File("~/PDF/project.pdf", "application/pdf", "project.pdf");
        }

        [HttpPost]
        public ActionResult SelectFCUOption(long ProjectID, long UnitID, string ModelID)
        {
            OptionDataRequest optionsRequest = new OptionDataRequest();
            
            try
            {
                foreach (string item in ModelID.Split(','))
                {
                    string modelId = DbAccess.GetModelData(conn, item, UnitID.ToString());
                    int selCount = objGJWFCU.SelectFCUoption(conn, ProjectID, UnitID, Convert.ToInt64(modelId));
                    if (selCount != 0)
                    {
                        optionsRequest.Status = "Failure";
                        optionsRequest.ErrorMessage = "";
                    }
                    
                }
                
            }
            catch (NullReferenceException ex)
            {

                optionsRequest.Status = "Failure";
                optionsRequest.ErrorMessage = ex.Message;
            }
            catch (Exception ex)
            {

                optionsRequest.Status = "Failure";
                optionsRequest.ErrorMessage = ex.Message;
            }

            return Json(optionsRequest);
        }

        [HttpPost]
        public FileResult ExportExcel(string GridHtml, string ProjName)
        {
            return File(Encoding.ASCII.GetBytes(GridHtml), "application/vnd.ms-excel", ProjName+"_Schedule_Export.xls");
        }

        [HttpPost]
        public ActionResult AddRecord(string name,string projectID, string typeName)
        {
            OptionDataRequest optionsRequest = new OptionDataRequest();
            try
            {
                if (typeName == "Add Project")
                {
                    objGJWFCU.addProject(conn,UserID, name);
                    optionsRequest.ErrorMessage = "Project added successfully";
                }
                else
                {
                    objGJWFCU.AddUnit(conn, Convert.ToInt64(projectID), name);
                    optionsRequest.ErrorMessage = "Unit added successfully";
                }
                optionsRequest.Status = "Success";
            }
            catch (Exception ex)
            {

                optionsRequest.Status = "Failure";
                optionsRequest.ErrorMessage = ex.Message;
            }
            
            return Json(optionsRequest);
        }

        [HttpPost]
        public ActionResult DeleteSelection(string unitId, string modelId)
        {
            OptionDataRequest optionsRequest = new OptionDataRequest();
            DbAccess.DeleteSelection(conn,unitId, modelId);
            optionsRequest.Status = "Success";
            optionsRequest.ErrorMessage = "";
            return Json(optionsRequest);
        }

        [HttpPost]
        public ActionResult UpdateProjectSettings(string name, string projectID, string typeName)
        {
            OptionDataRequest optionsRequest = new OptionDataRequest();
            if (typeName == "Add Project")
            {
                objGJWFCU.addProject(conn, 0, name);
            }
            else
            {
                objGJWFCU.AddUnit(conn, Convert.ToInt64(projectID), name);
            }
            optionsRequest.Status = "Success";
            optionsRequest.ErrorMessage = "";
            return Json(optionsRequest);
        }

        [HttpPost]
        public ActionResult UpdateProjectDetails(string projectId,string minLADBColling, string minLADBHeating, string minCapacitySensible, string maxCapacitySensible,string pressMarginMinPC,string coilFVmin, string coilFVMax)
        {
            OptionDataRequest optionsRequest = new OptionDataRequest();
            DbAccess.UpdateProjectDetails(conn, projectId, minLADBColling,  minLADBHeating,  minCapacitySensible,  maxCapacitySensible,  pressMarginMinPC,  coilFVmin,  coilFVMax);
            optionsRequest.Status = "Success";
            optionsRequest.ErrorMessage = "";
            return Json(optionsRequest);
        }

        [HttpPost]
        public ActionResult SaveProjectDuplicateUnit(string name, string projectID, string unitId, string typeName)
        {
            string msg = "";
            OptionDataRequest optionsRequest = new OptionDataRequest();
            try
            {
                if (typeName == "Save As Project")
                {
                    objGJWFCU.DuplicateProject(conn, Convert.ToInt64(projectID), name);
                    msg = "Project saved as " + name;
                }
                else
                {
                    objGJWFCU.DuplicateUnit(conn, Convert.ToInt64(projectID), Convert.ToInt64(unitId), name);
                    msg = "Unit saved as " + name;
                }
                optionsRequest.Status = "Success";
                optionsRequest.ErrorMessage = msg;
            }
            catch (Exception ex)
            {
                optionsRequest.Status = "Failure";
                optionsRequest.ErrorMessage = ex.Message;
            }
            
           
            return Json(optionsRequest);
        }

        [HttpPost]
        public ActionResult SaveSelection(string? selectionRequestModel)
        {
            string msg = "";
            OptionDataRequest optionsRequest = new OptionDataRequest();
            try
            {
                if (selectionRequestModel != null)
                {
                    SelectionRequestModel result = JsonConvert.DeserializeObject<SelectionRequestModel>(selectionRequestModel);
                    if (result != null)
                       DbAccess.UpdateProjectSelectionDetails(conn, result);
                }
                msg = "Scheduling is saved successfully";
                optionsRequest.Status = "Success";
                optionsRequest.ErrorMessage = msg;
            }
            catch (Exception ex)
            {
                optionsRequest.Status = "Failure";
                optionsRequest.ErrorMessage = ex.Message;
            }


            return Json(optionsRequest);
        }

        [HttpPost]
        public ActionResult PrintSchedule(string selectionRequestModel)
        {
            DataTable dtOption = null;
            SelectionRequestModel result = new SelectionRequestModel();
            if (selectionRequestModel != null)
            {
               
                result = JsonConvert.DeserializeObject<SelectionRequestModel>(selectionRequestModel);
                
            }
            if(result!=null)
            {
                result.DatedOn = DateTime.Now.ToString("MMMM dd, yyyy");
                result.CoilData = new List<ColiDataModel>();
                result.CoilHeatData = new List<ColiDataModel>();
                result.FanData = new List<FanDataModel>();
                result.AttributeData = new List<AttributeModel>();
                result.AncillariesData = new List<AncillariesModel>();
                result.ModelImageData = new List<ModelImages>();
                result.UnitData = new List<UnitPrintResult>();
                foreach (var item in result.selectionModelRequestModel)
                {
                    dtOption = DbAccess.GetCoilSelectionData(conn, item.UnitModelID.ToString(), item.UnitId.ToString());
                    //Add Coil Data
                    List<ColiDataModel> options = DbAccess.CreateListFromTable<ColiDataModel>(dtOption);
                    if(options!=null)
                    {
                        if (options[0] != null)
                            result.CoilData.Add(options[0]);
                        if (options[1] != null)
                            result.CoilHeatData.Add(options[1]);
                    }

                    //Add Fan Data
                    dtOption = DbAccess.GetFanSelectionData(conn, item.UnitModelID.ToString(), item.UnitId.ToString());
                    List<FanDataModel> optionsFanData = DbAccess.CreateListFromTable<FanDataModel>(dtOption);
                    if (optionsFanData != null)
                    {
                        if (optionsFanData[0] != null)
                        {
                            DataTable curveData = objGJWFCU.GetFanSysCurve(conn, item.UnitId, item.UnitModelID);
                            List<DataPoint> fanPress = new List<DataPoint>();
                            List<DataPoint> sysPress = new List<DataPoint>();
                            List<FanCurveData> fanCurveData = DbAccess.CreateListFromTable<FanCurveData>(curveData).OrderBy(p => p.Airflow).Distinct().ToList();
                            int cnt = 100;
                            foreach (var itemCurve in fanCurveData)
                            {
                                fanPress.Add(new DataPoint(itemCurve.Airflow, Convert.ToDouble(itemCurve.FanPress), "FanPress"));
                                sysPress.Add(new DataPoint(itemCurve.Airflow, Convert.ToDouble(itemCurve.SysPress), "SysPress"));
                                cnt = cnt + 100;
                            }
                            optionsFanData[0].FanPress = JsonConvert.SerializeObject(fanPress);
                            optionsFanData[0].SysPress = JsonConvert.SerializeObject(sysPress);
                            result.FanData.Add(optionsFanData[0]);
                        }
                    }

                    //Add Attribute Data
                    dtOption = DbAccess.GetAttributeSelectionData(conn, item.UnitModelID.ToString(), item.UnitId.ToString());
                    List<AttributeModel> optionsAttribute = DbAccess.CreateListFromTable<AttributeModel>(dtOption);
                    if (optionsAttribute != null)
                    {
                        if (optionsAttribute[0] != null)
                            result.AttributeData.Add(optionsAttribute[0]);
                    }

                    //Add Ancillaries Data
                    dtOption = DbAccess.GetAncillariesSelectionData(conn, item.UnitModelID.ToString(), item.UnitId.ToString());
                    List<AncillariesModel> optionsAncillaries = DbAccess.CreateListFromTable<AncillariesModel>(dtOption);
                    if (optionsAncillaries != null)
                    {
                        if (optionsAncillaries[0] != null)
                            result.AncillariesData.Add(optionsAncillaries[0]);
                    }

                    //Add Model Images
                    dtOption = DbAccess.GetSelectionModelImages(conn, item.UnitModelID.ToString(), item.UnitId.ToString());
                    List<ModelImages> optionsModelImages = DbAccess.CreateListFromTable<ModelImages>(dtOption);
                    if (optionsModelImages != null)
                    {
                        if (optionsModelImages[0] != null)
                            result.ModelImageData.Add(optionsModelImages[0]);
                    }

                    //Add Unit Data
                    dtOption = DbAccess.GetPrintUnitById(conn, item.UnitId.ToString());
                    List<UnitPrintResult> optionsUnit = DbAccess.CreateListFromTable<UnitPrintResult>(dtOption);
                    if (optionsUnit != null)
                    {
                        if (optionsUnit[0] != null)
                            result.UnitData.Add(optionsUnit[0]);
                    }
                }

            }
            return PartialView("_schedulingQuote", result);
        }

        [HttpPost]
        public ActionResult ExportExcelHTML(string selectionRequestModel)
        {
            DataTable dtOption = null;
            SelectionRequestModel result = new SelectionRequestModel();
            if (selectionRequestModel != null)
            {

                result = JsonConvert.DeserializeObject<SelectionRequestModel>(selectionRequestModel);

            }
            if (result != null)
            {
                result.DatedOn = DateTime.Now.ToString("MMMM dd, yyyy");
                result.CoilData = new List<ColiDataModel>();
                result.CoilHeatData = new List<ColiDataModel>();
                result.FanData = new List<FanDataModel>();
                result.AttributeData = new List<AttributeModel>();
                result.AncillariesData = new List<AncillariesModel>();
                result.ModelImageData = new List<ModelImages>();
                result.UnitData = new List<UnitPrintResult>();
                foreach (var item in result.selectionModelRequestModel)
                {
                    dtOption = DbAccess.GetCoilSelectionData(conn, item.UnitModelID.ToString(), item.UnitId.ToString());
                    //Add Coil Data
                    List<ColiDataModel> options = DbAccess.CreateListFromTable<ColiDataModel>(dtOption);
                    if (options != null)
                    {
                        if (options[0] != null)
                            result.CoilData.Add(options[0]);
                        if (options[1] != null)
                            result.CoilHeatData.Add(options[1]);
                    }

                    //Add Fan Data
                    dtOption = DbAccess.GetFanSelectionData(conn, item.UnitModelID.ToString(), item.UnitId.ToString());
                    List<FanDataModel> optionsFanData = DbAccess.CreateListFromTable<FanDataModel>(dtOption);
                    if (optionsFanData != null)
                    {
                        if (optionsFanData[0] != null)
                        {
                            DataTable curveData = objGJWFCU.GetFanSysCurve(conn, item.UnitId, item.UnitModelID);
                            List<DataPoint> fanPress = new List<DataPoint>();
                            List<DataPoint> sysPress = new List<DataPoint>();
                            List<FanCurveData> fanCurveData = DbAccess.CreateListFromTable<FanCurveData>(curveData).OrderBy(p => p.Airflow).Distinct().ToList();
                            int cnt = 100;
                            foreach (var itemCurve in fanCurveData)
                            {
                                fanPress.Add(new DataPoint(itemCurve.Airflow, Convert.ToDouble(itemCurve.FanPress), "FanPress"));
                                sysPress.Add(new DataPoint(itemCurve.Airflow, Convert.ToDouble(itemCurve.SysPress), "SysPress"));
                                cnt = cnt + 100;
                            }
                            optionsFanData[0].FanPress = JsonConvert.SerializeObject(fanPress);
                            optionsFanData[0].SysPress = JsonConvert.SerializeObject(sysPress);
                            result.FanData.Add(optionsFanData[0]);
                        }
                    }

                    //Add Attribute Data
                    dtOption = DbAccess.GetAttributeSelectionData(conn, item.UnitModelID.ToString(), item.UnitId.ToString());
                    List<AttributeModel> optionsAttribute = DbAccess.CreateListFromTable<AttributeModel>(dtOption);
                    if (optionsAttribute != null)
                    {
                        if (optionsAttribute[0] != null)
                            result.AttributeData.Add(optionsAttribute[0]);
                    }

                    //Add Ancillaries Data
                    dtOption = DbAccess.GetAncillariesSelectionData(conn, item.UnitModelID.ToString(), item.UnitId.ToString());
                    List<AncillariesModel> optionsAncillaries = DbAccess.CreateListFromTable<AncillariesModel>(dtOption);
                    if (optionsAncillaries != null)
                    {
                        if (optionsAncillaries[0] != null)
                            result.AncillariesData.Add(optionsAncillaries[0]);
                    }

                    //Add Model Images
                    dtOption = DbAccess.GetSelectionModelImages(conn, item.UnitModelID.ToString(), item.UnitId.ToString());
                    List<ModelImages> optionsModelImages = DbAccess.CreateListFromTable<ModelImages>(dtOption);
                    if (optionsModelImages != null)
                    {
                        if (optionsModelImages[0] != null)
                            result.ModelImageData.Add(optionsModelImages[0]);
                    }

                    //Add Unit Data
                    dtOption = DbAccess.GetPrintUnitById(conn, item.UnitId.ToString());
                    List<UnitPrintResult> optionsUnit = DbAccess.CreateListFromTable<UnitPrintResult>(dtOption);
                    if (optionsUnit != null)
                    {
                        if (optionsUnit[0] != null)
                            result.UnitData.Add(optionsUnit[0]);
                    }
                }

            }
            return PartialView("_excelImport", result);
        }
    }
}
