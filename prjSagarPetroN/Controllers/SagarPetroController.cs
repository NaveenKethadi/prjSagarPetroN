using ClosedXML.Excel;
using Newtonsoft.Json;
using prjSagarPetroN.Models;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Web.Configuration;
using System.Web.Mvc;
using System.Web.Script.Serialization;

namespace prjSagarPetroN.Controllers
{
    public class SagarPetroController : Controller
    {
        // GET: SagarPetro
        string error = string.Empty;
        private static readonly string baseUrl = WebConfigurationManager.AppSettings["Server_API_IP"];
        public static string ProdPlan => baseUrl + "/Transactions/Vouchers/Production Planning";//5633 
        public static string PurInd => baseUrl + "/Transactions/Vouchers/Purchase Indent";//7937 PurInd
        ClsDataAcceslayer obj = new ClsDataAcceslayer();
        public ActionResult Index(int companyId)
        {
            ViewBag.cid = companyId;
            //ViewBag.Items = GetItems(companyId);
            ViewBag.Itemgroup = GetItemsLastGroup(companyId);
            ViewBag.wh = GetWarehouse(companyId);
            ViewBag.Parent = GetParent(companyId);
            return View();
        }
        public ActionResult GetItems(int selectedId, int companyId)//IEnumerable<SelectListItem>
        {
            //string Message = "";
            Log("Selected Item Group Id :" + selectedId);
            //int parentid;
            List<SelectListItem> Itemlist = new List<SelectListItem>();
            IEnumerable<SelectListItem> inchargelst = new List<SelectListItem>();
            if (selectedId == 412 || selectedId == 5)
            {
                inchargelst = GetIncharge(companyId, selectedId);// IEnumerable<SelectListItem>
            }
            string splistQry = $@"select iMasterId,sName from fCore_GetProductTreeSequence({selectedId},0) where bGroup=0 and iMasterId>0";
            DataSet ds = ClsDataAcceslayer.GetData1(splistQry, companyId, ref error);
            if (ds != null && ds.Tables[0].Rows.Count > 0)//&& inchargelst.Count()>0
            {
                for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                {
                    Itemlist.Add(new SelectListItem()
                    {
                        Value = ds.Tables[0].Rows[i]["iMasterId"].ToString(),
                        Text = ds.Tables[0].Rows[i]["sName"].ToString()
                    });
                }

                return Json(new { status = true, Itemlist, inchargelst });
            }

            else
            {
                Log($"No items are available for the selected item group(item ID :{selectedId})");
                return Json(new { status = false, Message = "No items are available for the selected item group" });
            }
            //return new SelectList(Itemlist.AsEnumerable(), "Value", "Text");

        }
        public IEnumerable<SelectListItem> GetIncharge(int cid, int selid)
        {
            int pid = 0;
            List<SelectListItem> inchrgelist = new List<SelectListItem>();
            inchrgelist.Add(new SelectListItem { Value = "0", Text = "--select--" });
            if (selid == 412) pid = 93;
            if (selid == 5) pid = 92;
            string incQry = $@"select iMasterId,sName from vmCore_Account where  iParentId={pid} and iStatus<>5";// --FG Incharge
            DataSet ds = ClsDataAcceslayer.GetData1(incQry, cid, ref error);
            if (ds != null && ds.Tables[0].Rows.Count > 0)
            {
                for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                {
                    inchrgelist.Add(new SelectListItem()
                    {
                        Value = ds.Tables[0].Rows[i]["iMasterId"].ToString(),
                        Text = ds.Tables[0].Rows[i]["sName"].ToString()
                    });
                }
            }
            return inchrgelist.AsEnumerable();
        }
        public IEnumerable<SelectListItem> GetItemsLastGroup(int companyId)
        {
            List<SelectListItem> Itemgrplist = new List<SelectListItem>();
            Itemgrplist.Add(new SelectListItem { Value = "0", Text = "--select--" });
            string splistQry = $@"select iMasterId,sName from mCore_Product where bGroup=1 and iStatus<>5 and sCode in ('II','PM','FG','RM/001','RM/002')";//,'CA'
            DataSet ds = ClsDataAcceslayer.GetData1(splistQry, companyId, ref error);
            for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
            {
                Itemgrplist.Add(new SelectListItem()
                {
                    Value = ds.Tables[0].Rows[i]["iMasterId"].ToString(),
                    Text = ds.Tables[0].Rows[i]["sName"].ToString()
                });
            }

            return new SelectList(Itemgrplist.AsEnumerable(), "Value", "Text");
        }
        public IEnumerable<SelectListItem> GetWarehouse(int companyId)
        {
            List<SelectListItem> Whlist = new List<SelectListItem>();
            Whlist.Add(new SelectListItem { Value = "0", Text = "--select--" });
            string splistQry = $@"select iMasterId,sName from mCore_Warehouse where iStatus<>5 and iMasterId<>0";
            DataSet ds = ClsDataAcceslayer.GetData1(splistQry, companyId, ref error);
            for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
            {
                Whlist.Add(new SelectListItem()
                {
                    Value = ds.Tables[0].Rows[i]["iMasterId"].ToString(),
                    Text = ds.Tables[0].Rows[i]["sName"].ToString()
                });
            }

            return new SelectList(Whlist.AsEnumerable(), "Value", "Text");
        }
        //GetItemGroups
        public ActionResult GetItemGroups(int selectedId, int companyId)
        {
            JavaScriptSerializer jss = new JavaScriptSerializer();
            string Str2 = string.Empty;
            List<ClsPrepaymentsMstr> PrePaymentMstrs2 = new List<ClsPrepaymentsMstr>();
            string error = string.Empty;
            Str2 = $@"select p.iMasterId,sName,ISNULL(iParentId,0) iParentId from mCore_Product p
                     join mCore_ProductTreeDetails mp on p.iMasterId = mp.iMasterId where p.iMasterId > 0
                     and p.iMasterId in (select iMasterId from fCore_GetProductTreeSequence({selectedId},0)) 
                    and iParentId<>0
                    order by sName";
            DataSet ds2 = ClsDataAcceslayer.GetData1(Str2, Convert.ToInt32(companyId), ref error);
            for (int i = 0; i < ds2.Tables[0].Rows.Count; i++)
            {
                PrePaymentMstrs2.Add(new ClsPrepaymentsMstr { id = ds2.Tables[0].Rows[i]["iMasterId"].ToString(), pid = ds2.Tables[0].Rows[i]["iParentId"].ToString(), name = ds2.Tables[0].Rows[i]["sName"].ToString() });
            }

            //return View(PrePaymentMstrs2);
            var jsonResult = Json(new { status = true, PrePaymentMstrs2 }, JsonRequestBehavior.AllowGet);
            jsonResult.MaxJsonLength = int.MaxValue;
            return jsonResult;
        }
        public ActionResult CheckReturnType(string selectedIds, int companyId)
        {
            try
            {
                bool status1 = true;
                //query to get only sub groups from the selected selectedIds
                string q = $@"select distinct iParentId  as subgrpid
					 from vmCore_Product 
					 where iMasterId in ({selectedIds})
					 and bgroup=0";
                DataSet ds = ClsDataAcceslayer.GetData1(q, companyId, ref error);
                string subgrpids = "";string diffrepr1 = "";
                List<int> a = new List<int>();
                List<SubGrpReptcheckModel> subgrprptlst = new List<SubGrpReptcheckModel>();
                if (ds != null && ds.Tables[0].Rows.Count > 0)
                {
                    for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                    {
                        a.Add(Convert.ToInt32(ds.Tables[0].Rows[i]["subgrpid"]));
                    }
                }
                if (a.Count > 0)
                {
                    subgrpids = string.Join(",", a);
                    string q1 = $@"select mp.sName,mup.ReportType from mCore_Product mp join 
				                muCore_Product mup on mp.iMasterId=mup.iMasterId 				
				                where mp.iMasterId in ({subgrpids})";
                    DataSet ds1 = ClsDataAcceslayer.GetData1(q1, companyId, ref error);
                    if (ds1 != null && ds1.Tables[0].Rows.Count > 0)
                    {
                        for (int i = 0; i < ds1.Tables[0].Rows.Count; i++)
                        {
                            SubGrpReptcheckModel subgrprpt = new SubGrpReptcheckModel();
                            subgrprpt.Id = Convert.ToInt32(ds1.Tables[0].Rows[i]["ReportType"]);
                            subgrprpt.Name = ds1.Tables[0].Rows[i]["sName"].ToString();
                            subgrprptlst.Add(subgrprpt);
                        }

                    }
                    var groups = subgrprptlst.GroupBy(x => x.Id).ToList();
                    if (groups.Count == 1)
                    {
                        Log($"All Sub group's report types are the same-{groups.FirstOrDefault().Key}");
                        var rt = groups.FirstOrDefault().Key;var rt1 = 0; string msg = "";
                        IEnumerable<SelectListItem> inchargelst = new List<SelectListItem>();
                        if (rt == 2 || rt == 3)
                        {
                            if (rt == 2) rt1 = 412;
                            if (rt == 3) rt1 = 5;
                            inchargelst = GetIncharge(companyId, rt1);// IEnumerable<SelectListItem>
                                                                      //                    
                        }
                        else if (rt == 1)
                        {
                            status1 = false;
                            // string msg = $"Item:{groups.ElementAt(0).FirstOrDefault().Name}'s report type is not defined";
                            msg = $"Report type of Item -'{groups.ElementAt(0).FirstOrDefault().Name}' is not defined";
                            // return Json(new { status = false, Message = msg });
                        }
                       // else
                          //  rt1 = rt;
                        return Json(new { status = status1, rt, inchargelst , Message = msg });
                    }
                    else
                    {
                        StringBuilder diffrepr = new StringBuilder();
                        Log("Items with different report types:");
                        foreach (var group in groups)
                        {
                            foreach (var item in group)
                            {
                                string reptp = "";
                                if (item.Id == 1) reptp = "Report Type not defined";
                                if (item.Id == 2) reptp = "FG";if (item.Id == 3) reptp = "SFG";if (item.Id == 4) reptp = "Base Oil";
                                if (item.Id == 5) reptp = "Additives"; if (item.Id == 6) reptp = "Packing Material";
                                diffrepr.Append($"{item.Name}:{reptp},");
                                Log($"Name: {item.Name}, ReportType: {item.Id}");
                            }
                        }
                        if (diffrepr.Length > 0)
                        {
                            diffrepr.Length--; // Remove the last character, which is the comma
                            diffrepr1 = "Items with different report types found\n";
                            diffrepr1 += $"{diffrepr}";
                        }
                        return Json(new { status = false, Message = diffrepr1 });
                    }
                }
                return Json(0);
            }
            catch (Exception a)
            {
                return Json(new { status = false, Message = a.Message });
            }
        }
        public IEnumerable<SelectListItem> GetParent(int companyId)
        {
            List<SelectListItem> Whlist = new List<SelectListItem>();
            Whlist.Add(new SelectListItem { Value = "0", Text = "--select--" });
            string splistQry = $@"select distinct p.sName,p.iMasterId from  dbo.fCore_GetProductByLevel(0) l
                join mCore_Product p  on l.iParentId = p.iMasterId and iStatus<>5
                where bGroup = 1
                order by p.iMasterId";
            DataSet ds = ClsDataAcceslayer.GetData1(splistQry, companyId, ref error);
            for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
            {
                Whlist.Add(new SelectListItem()
                {
                    Value = ds.Tables[0].Rows[i]["iMasterId"].ToString(),
                    Text = ds.Tables[0].Rows[i]["sName"].ToString()
                });
            }

            return new SelectList(Whlist.AsEnumerable(), "Value", "Text");
        }
        // public ActionResult GetFGItemsonStocklevel(List<int> selectedValues, int companyId, string dt, int GroupId)
        public ActionResult GetFGItemsonStocklevel(string selected, int companyId, string dt)//, int GroupId
        {
            try
            {
                Log("GetFGItemsonStocklevel Method");
               // Log("Selected Group:" + GroupId + ",companyId:" + companyId);
                //string commaseperateditems = "";
                List<ItemShowModel> itemshowmodellist1 = new List<ItemShowModel>();
                //foreach (var item in selectedValues)
                //{
                //    commaseperateditems = commaseperateditems + item.ToString() + ",";
                //}
                Log("Selected Items :" + selected);
                List<int> selectedValues = selected.Split(',')
                                  .Select(int.Parse)
                                  .ToList();
                foreach (var item in selectedValues)
                {
                    int avgsalecnsumdays = 0;
                    int reorderdangerdays = 0;
                    int safetystockdays = 0;
                    decimal consumeval = 0;
                    decimal reorderval = 0;
                    decimal safetystockval = 0;
                    decimal Astockval = 0;

                    string qgetdays = $@"select AvgSalesConsumptioninDay,ReOrderDangerLevelinDays,SafetyStockLevelinDays 
                                from muCore_Product  where iMasterId={item}";
                    Log("" + qgetdays);
                    DataSet dsgetdays = ClsDataAcceslayer.GetData1(qgetdays, companyId, ref error);
                    Log("query result count:" + dsgetdays.Tables[0].Rows.Count);
                    if (dsgetdays != null && dsgetdays.Tables[0].Rows.Count > 0)
                    {
                        avgsalecnsumdays = Convert.ToInt32(dsgetdays.Tables[0].Rows[0]["AvgSalesConsumptioninDay"]);
                        reorderdangerdays = Convert.ToInt32(dsgetdays.Tables[0].Rows[0]["ReOrderDangerLevelinDays"]);
                        safetystockdays = Convert.ToInt32(dsgetdays.Tables[0].Rows[0]["SafetyStockLevelinDays"]);//abs  
                        if (avgsalecnsumdays == 0) { return Json(new { status = false, Message = "Please Enter valid Avg Sales/Consumption in Day(s) in Item master" }); }
                        Log("For Item :" + item + ", AvgSalesConsumptioninDays :" + avgsalecnsumdays + ",ReOrderDangerLevelinDays :" + reorderdangerdays + ",SafetyStockLevelinDays :" + safetystockdays);
                        if (Convert.ToInt32(dsgetdays.Tables[0].Rows[0]["AvgSalesConsumptioninDay"]) != 0)
                        {
                            string qgetvalue = $@"select isnull(cast(abs(sum(fQuantity)) as decimal(18,2)),0) fq
                                from tCore_Header_0 h 
                                join tCore_Data_0 d on d.iHeaderId = h.iHeaderId
                                join tCore_Indta_0 ind on ind.iBodyId = d.iBodyId
                                where fQuantity < 0 and  iDate >= dbo.DateToInt(DATEADD(day,-{Convert.ToInt32(dsgetdays.Tables[0].Rows[0]["AvgSalesConsumptioninDay"])}, GETDATE())) 
                                and ind.iProduct={item} and bUpdateStocks=1";
                            Log("" + qgetvalue);
                            DataSet dsgetvalue = ClsDataAcceslayer.GetData1(qgetvalue, companyId, ref error);
                            Log("Query result count:" + dsgetvalue.Tables[0].Rows.Count);
                            if (dsgetvalue != null && dsgetvalue.Tables[0].Rows.Count > 0)
                            {
                                consumeval = Convert.ToDecimal(dsgetvalue.Tables[0].Rows[0]["fq"]);
                                reorderval = (consumeval / avgsalecnsumdays) * reorderdangerdays;
                                safetystockval = (consumeval / avgsalecnsumdays) * safetystockdays;
                            }
                            string qgetstock = $@"select isnull(sum(fQiss+fQrec),0) stock from vCore_ibals_0 where iProduct={item}";
                            Log("stock query:" + qgetstock);
                            DataSet dsgetstock = ClsDataAcceslayer.GetData1(qgetstock, companyId, ref error);
                            if (dsgetstock != null && dsgetstock.Tables[0].Rows.Count > 0)
                            {
                                Astockval = Convert.ToDecimal(dsgetstock.Tables[0].Rows[0]["stock"]);
                                Log("Actual Stock:" + Astockval);
                            }
                            Log("safetylevel Stock:" + safetystockval);
                            if (Astockval < safetystockval)
                            {

                                //                   string qitemdetails = $@"select Item,Itemid,sCode,UnitId,DefaultBaseUnit,Description,isnull(sVoucherNo,'') sVoucherNo,isnull(sProdOrderNo,'') sProdOrderNo,isnull(duedate,'') duedate from
                                //                       (select mp.sName Item,mp.iMasterId Itemid,mp.sCode,mu.sName DefaultBaseUnit,
                                //                       mu.iMasterId UnitId,sDescription Description 
                                //                       from mCore_Product mp
                                //                       join muCore_Product_Units mpu on mpu.iMasterId=mp.iMasterId
                                //                       join mCore_Units mu on mu.iMasterId=iDefaultBaseUnit
                                //                       join muCore_Product mup on mup.iMasterId=mp.iMasterId
                                //                       ) k
                                //left  join
                                //                       (select top 1 sVoucherNo,iItem,sProdOrderNo,convert(varchar,dbo.IntToDate(DueDate),103) duedate from tCore_Header_0 h join tCore_Data_0 d on h.iHeaderId=d.iHeaderId
                                //                       join tCore_Indta_0 ind on ind.iBodyId=d.iBodyId 
                                //                       join tCore_Data5633_0 dv on dv.iBodyId=d.iBodyId
                                //                       join tMrp_ProdOrderBody_0 pob on pob.iItem=ind.iProduct
                                //                       join tMrp_ProdOrder_0  po on po.iProdOrderId=pob.iProdOrderId
                                //                       where  iVoucherType=5633 and pob.fQuantity={-(Astockval - safetystockval)} order by 1 desc) l on k.Itemid=l.iItem
                                //                       where Itemid={item}";
                                string qitemdetails = $@"select Item,Itemid,sCode,UnitId,DefaultBaseUnit,Description,
                                isnull(sVoucherNo,'') sVoucherNo,isnull(ProductionOrderNo_,'') ProductionOrderNo_,isnull(duedate,'') duedate from
                                (select mp.sName Item,mp.iMasterId Itemid,mp.sCode,mu.sName DefaultBaseUnit,
                                mu.iMasterId UnitId,sDescription Description 
                                from mCore_Product mp
                                join muCore_Product_Units mpu on mpu.iMasterId=mp.iMasterId
                                join mCore_Units mu on mu.iMasterId=iDefaultBaseUnit
                                join muCore_Product mup on mup.iMasterId=mp.iMasterId
                                ) k
	                                left  join
                                (select top 1 sVoucherNo,iProduct,ProductionOrderNo_,convert(varchar,dbo.IntToDate(DueDate),103) duedate from tCore_Header_0 h join tCore_Data_0 d on h.iHeaderId=d.iHeaderId
                                join tCore_Indta_0 ind on ind.iBodyId=d.iBodyId 
                                join tCore_Data5633_0 dv on dv.iBodyId=d.iBodyId
                                --join tMrp_ProdOrderBody_0 pob on pob.iItem=ind.iProduct
                                --join tMrp_ProdOrder_0  po on po.iProdOrderId=pob.iProdOrderId
                                where  iVoucherType=5633 order by LEN(sVoucherNo) desc,sVoucherNo desc) l on k.Itemid=l.iProduct
                                where Itemid={item}";
                                Log("" + qitemdetails);
                                DataSet dsitemdetails = ClsDataAcceslayer.GetData1(qitemdetails, companyId, ref error);
                                Log("query result count:" + dsitemdetails.Tables[0].Rows.Count);
                                if (dsitemdetails != null && dsitemdetails.Tables[0].Rows.Count > 0)
                                {
                                    ItemShowModel itemshowmodel = new ItemShowModel();
                                    itemshowmodel.Item = dsitemdetails.Tables[0].Rows[0]["Item"].ToString();
                                    itemshowmodel.Itemid = Convert.ToInt32(dsitemdetails.Tables[0].Rows[0]["Itemid"]);
                                    itemshowmodel.UnitId = Convert.ToInt32(dsitemdetails.Tables[0].Rows[0]["UnitId"]);
                                    itemshowmodel.Units = dsitemdetails.Tables[0].Rows[0]["DefaultBaseUnit"].ToString();
                                    itemshowmodel.Description = dsitemdetails.Tables[0].Rows[0]["Description"].ToString();
                                    itemshowmodel.AvailableStockQty = Astockval;
                                    itemshowmodel.SafetyLevelQty = safetystockval;
                                    itemshowmodel.Difference = Astockval - safetystockval;
                                    itemshowmodel.QtytobeProduced = -(Astockval - safetystockval);
                                    if (dsitemdetails.Tables[0].Rows[0]["sVoucherNo"].ToString() != "")
                                        itemshowmodel.ProductionPlanningStatus = "Raised";
                                    else
                                        itemshowmodel.ProductionPlanningStatus = "Not Raised";
                                    itemshowmodel.ProductionPlanningDocNo = dsitemdetails.Tables[0].Rows[0]["sVoucherNo"].ToString();
                                    itemshowmodel.DueDate = dsitemdetails.Tables[0].Rows[0]["duedate"].ToString();
                                    itemshowmodel.ItemCode = dsitemdetails.Tables[0].Rows[0]["sCode"].ToString();
                                    itemshowmodel.sVoucherNo = dsitemdetails.Tables[0].Rows[0]["sVoucherNo"].ToString();
                                    itemshowmodellist1.Add(itemshowmodel);
                                }
                            }
                        }
                        else { Log("Please Enter valid Avg Sales/Consumption in Day(s) in Item master for item:" + item); }
                    }
                }
                var itemshowmodellist = itemshowmodellist1.OrderBy(_ => _.QtytobeProduced);
                return Json(new { status = true, itemshowmodellist });
            }
            catch (Exception x)
            {
                return Json(new { status = false, Message = x.Message });
            }
        }
        //GetSFGItemsonStocklevel
        //  public ActionResult GetSFGItemsonStocklevel(List<int> selectedValues, int companyId, string dt, int GroupId)
        public ActionResult GetSFGItemsonStocklevel(string selected, int companyId, string dt)
        {
            try
            {
                Log("GetSFGItemsonStocklevel Method");
             //   Log("Selected Group:" + GroupId + ",companyId:" + companyId);
                //string commaseperateditems = "";
                List<ItemShowModel> itemshowmodellist1 = new List<ItemShowModel>();
                //foreach (var item in selectedValues)
                //{
                //    commaseperateditems = commaseperateditems + item.ToString() + ",";
                //}
                //Log("Selected Items :" + commaseperateditems.Trim(','));
                List<int> selectedValues = selected.Split(',')
                                   .Select(int.Parse)
                                   .ToList();
                foreach (var item in selectedValues)
                {
                    Log("Item Id:" + item);
                    int avgsalecnsumdays = 0;
                    int reorderdangerdays = 0;
                    int safetystockdays = 0;
                    decimal consumeval = 0;
                    decimal reorderval = 0;
                    decimal safetystockval = 0;
                    decimal Astockval = 0;

                    string qgetdays = $@"select AvgSalesConsumptioninDay,ReOrderDangerLevelinDays,SafetyStockLevelinDays 
                                from muCore_Product  where iMasterId={item}";
                    Log("" + qgetdays);
                    DataSet dsgetdays = ClsDataAcceslayer.GetData1(qgetdays, companyId, ref error);
                    Log("query result count:" + dsgetdays.Tables[0].Rows.Count);
                    if (dsgetdays != null && dsgetdays.Tables[0].Rows.Count > 0)
                    {
                        avgsalecnsumdays = Convert.ToInt32(dsgetdays.Tables[0].Rows[0]["AvgSalesConsumptioninDay"]);
                        reorderdangerdays = Convert.ToInt32(dsgetdays.Tables[0].Rows[0]["ReOrderDangerLevelinDays"]);
                        safetystockdays = Convert.ToInt32(dsgetdays.Tables[0].Rows[0]["SafetyStockLevelinDays"]);//abs
                        Log("For Item :" + item + ", AvgSalesConsumptioninDays :" + avgsalecnsumdays + ",ReOrderDangerLevelinDays :" + reorderdangerdays + ",SafetyStockLevelinDays :" + safetystockdays);
                        if (avgsalecnsumdays == 0)
                        {
                            //return Json(new { status = false, Message = "Please Enter valid Avg Sales/Consumption in Day(s) in Item master" });
                            Log("Please Enter valid Avg Sales/Consumption in Day(s) in Item master for item:" + item);
                        }
                        else
                        {
                            string qgetvalue = $@"select isnull(cast(abs(sum(fQuantity)) as decimal(18,2)),0) fq
                                from tCore_Header_0 h 
                                join tCore_Data_0 d on d.iHeaderId = h.iHeaderId
                                join tCore_Indta_0 ind on ind.iBodyId = d.iBodyId
                                where fQuantity < 0 and  iDate >= dbo.DateToInt(DATEADD(day,-{Convert.ToInt32(dsgetdays.Tables[0].Rows[0]["AvgSalesConsumptioninDay"])}, GETDATE())) 
                                and bUpdateStocks=1
                                and ind.iProduct={item} ";
                            Log("" + qgetvalue);
                            DataSet dsgetvalue = ClsDataAcceslayer.GetData1(qgetvalue, companyId, ref error);
                            Log("Query result count:" + dsgetvalue.Tables[0].Rows.Count);
                            if (dsgetvalue != null && dsgetvalue.Tables[0].Rows.Count > 0)
                            {
                                consumeval = Convert.ToDecimal(dsgetvalue.Tables[0].Rows[0]["fq"]);
                                reorderval = (consumeval / avgsalecnsumdays) * reorderdangerdays;
                                safetystockval = (consumeval / avgsalecnsumdays) * safetystockdays;
                            }
                            string qgetstock = $@"select isnull(sum(fQiss+fQrec),0) stock from vCore_ibals_0 where iProduct={item}";
                            Log("stock query:" + qgetstock);
                            DataSet dsgetstock = ClsDataAcceslayer.GetData1(qgetstock, companyId, ref error);
                            Log("Query result count:" + dsgetstock.Tables[0].Rows.Count);
                            if (dsgetstock != null && dsgetstock.Tables[0].Rows.Count > 0)
                            {
                                Astockval = Convert.ToDecimal(dsgetstock.Tables[0].Rows[0]["stock"]);
                                Log("Actual Stock:" + Astockval);
                            }
                            Log("safetystock Stock:" + safetystockval);
                            if (Astockval < safetystockval)
                            {
                                //string qitemdetails = $@"Select Item,a.iMasterId Itemid,a.UnitId,a.sCode,DefaultBaseUnit Units,isnull(Description,'') Description,
                                //    TankMaster,TankCapacity,ClosingStock,TopUpTankCapacity from
                                //    (select mp.sName Item,mp.sCode,mp.iMasterId,mu.sName DefaultBaseUnit,mu.iMasterId UnitId,sDescription Description,
                                //     iProductType,ProductType from mCore_Product mp
                                //    join muCore_Product_Units mpu on mpu.iMasterId=mp.iMasterId
                                //    join mCore_Units mu on mu.iMasterId=iDefaultBaseUnit
                                //    join muCore_Product mup on mup.iMasterId=mp.iMasterId) a join
                                //    (select w.sName TankMaster,iProduct,isnull(mw.Capacity,0) TankCapacity,sum(fQiss+fQrec) ClosingStock,
                                //    (mw.Capacity-sum(fQiss+fQrec))
                                //    TopUpTankCapacity from tCore_ibals_0  ib join mCore_Warehouse w on ib.iInvTag=w.iMasterId
                                //    join muCore_Warehouse mw on mw.iMasterId=w.iMasterId
                                //    group by w.sName,mw.Capacity,iProduct
                                //    having sum(fQiss+fQrec)>0) b on a.iMasterId=b.iProduct
                                //    where a.iMasterId={item}";
                                string qitemdetails = $@"Select Item,a.iMasterId Itemid,a.UnitId,a.sCode,DefaultBaseUnit Units,
isnull(Description,'') Description,isnull(TankMaster,'') TankMaster,isnull(Capacity,0) TankCapacity,isnull(ClosingStock,0) ClosingStock,
isnull(TopUpTankCapacity,0) TopUpTankCapacity,isnull(sVoucherNo,'') sVoucherNo,
isnull(ProductionOrderNo_,'') sProdOrderNo,isnull(duedate,'') duedate from
(select mp.sName Item,mp.sCode,mp.iMasterId,mu.sName DefaultBaseUnit,mu.iMasterId UnitId,sDescription Description,
iProductType from mCore_Product mp--ProductType
join muCore_Product_Units mpu on mpu.iMasterId=mp.iMasterId
join mCore_Units mu on mu.iMasterId=iDefaultBaseUnit
join muCore_Product mup on mup.iMasterId=mp.iMasterId) a 

left join
(select w.sName TankMaster,iProduct,Capacity,sum(fQiss+fQrec) ClosingStock,
(mw.Capacity-sum(fQiss+fQrec))
TopUpTankCapacity from vCore_ibals_0  ib join mCore_Warehouse w on ib.iInvTag=w.iMasterId
join muCore_Warehouse mw on mw.iMasterId=w.iMasterId
group by w.sName,mw.Capacity,iProduct
having sum(fQiss+fQrec)>0) b on a.iMasterId=b.iProduct
left  join

(select top 1 sVoucherNo,iProduct,ProductionOrderNo_,convert(varchar,dbo.IntToDate(DueDate),103) duedate from tCore_Header_0 h join tCore_Data_0 d on h.iHeaderId=d.iHeaderId
join tCore_Indta_0 ind on ind.iBodyId=d.iBodyId 
join tCore_Data5633_0 dv on dv.iBodyId=d.iBodyId
--join tMrp_ProdOrderBody_0 pob on pob.iItem=ind.iProduct
--join tMrp_ProdOrder_0  po on po.iProdOrderId=pob.iProdOrderId
where  iVoucherType=5633 --and TankTopUpCapacity
--and pob.fQuantity={-(Astockval - safetystockval)} 
order by len(sVoucherNo) desc,sVoucherNo desc) l on l.iProduct=a.iMasterId
where a.iMasterId={item}";//isnull(mw.Capacity,0) 
                                Log("" + qitemdetails);
                                DataSet dsitemdetails = ClsDataAcceslayer.GetData1(qitemdetails, companyId, ref error);
                                Log("query result count:" + dsitemdetails.Tables[0].Rows.Count);
                                if (dsitemdetails != null && dsitemdetails.Tables[0].Rows.Count > 0)
                                {
                                    for (int i = 0; i < dsitemdetails.Tables[0].Rows.Count; i++)
                                    {
                                        ItemShowModel itemshowmodel = new ItemShowModel();
                                        itemshowmodel.Item = dsitemdetails.Tables[0].Rows[i]["Item"].ToString();
                                        itemshowmodel.Units = dsitemdetails.Tables[0].Rows[i]["Units"].ToString();
                                        itemshowmodel.Description = dsitemdetails.Tables[0].Rows[i]["Description"].ToString();
                                        itemshowmodel.AvailableStockQty = Astockval;
                                        itemshowmodel.SafetyLevelQty = safetystockval;
                                        itemshowmodel.Difference = Astockval - safetystockval;
                                        itemshowmodel.QtytobeProduced = -(Astockval - safetystockval);
                                        itemshowmodel.TankMaster = dsitemdetails.Tables[0].Rows[i]["TankMaster"].ToString();
                                        itemshowmodel.TankCapacity = Convert.ToDecimal(dsitemdetails.Tables[0].Rows[i]["TankCapacity"]);
                                        itemshowmodel.ClosingStock = Convert.ToDecimal(dsitemdetails.Tables[0].Rows[i]["ClosingStock"]);
                                        itemshowmodel.TopUpTankCapacity = Convert.ToDecimal(dsitemdetails.Tables[0].Rows[i]["TopUpTankCapacity"]);
                                        if (dsitemdetails.Tables[0].Rows[0]["sVoucherNo"].ToString() != "")
                                            itemshowmodel.ProductionPlanningStatus = "Raised";
                                        else
                                            itemshowmodel.ProductionPlanningStatus = "Not Raised";
                                        itemshowmodel.ProductionPlanningDocNo = dsitemdetails.Tables[0].Rows[0]["sVoucherNo"].ToString();
                                        itemshowmodel.DueDate = dsitemdetails.Tables[0].Rows[0]["duedate"].ToString();
                                        itemshowmodel.ItemCode = dsitemdetails.Tables[0].Rows[0]["sCode"].ToString();
                                        itemshowmodel.Remarks = "";
                                        itemshowmodel.Itemid = Convert.ToInt32(dsitemdetails.Tables[0].Rows[i]["Itemid"]);
                                        itemshowmodel.UnitId = Convert.ToInt32(dsitemdetails.Tables[0].Rows[i]["UnitId"]);
                                        itemshowmodel.AvailableStockQty1 = Astockval;
                                        itemshowmodel.SafetyLevelQty1 = safetystockval;
                                        itemshowmodel.Difference1 = Astockval - safetystockval;
                                        itemshowmodel.QtytobeProduced1 = -(Astockval - safetystockval);
                                        itemshowmodel.TopUpTankCapacity1 = Convert.ToDecimal(dsitemdetails.Tables[0].Rows[i]["TopUpTankCapacity"]);
                                        itemshowmodel.sVoucherNo = dsitemdetails.Tables[0].Rows[0]["sVoucherNo"].ToString();
                                        itemshowmodellist1.Add(itemshowmodel);
                                    }
                                }
                            }
                        }
                    }
                }
                var itemshowmodellist = itemshowmodellist1.OrderBy(_ => _.QtytobeProduced);

                return Json(new { status = true, itemshowmodellist });
            }
            catch (Exception x)
            {
                return Json(new { status = false, Message = x.Message });
            }
        }
        public ActionResult GetRMBaseOils(string selected, int companyId, string dt)//, int GroupId
        {
            try
            {
                Log("GetRMBaseOils Method");
               // Log("Selected Group:" + GroupId + ",companyId:" + companyId);
               // string commaseperateditems = "";
                List<ItemShowModel> itemshowmodellist1 = new List<ItemShowModel>();
                List<int> selectedValues = selected.Split(',')
                                  .Select(int.Parse)
                                  .ToList();
                Log("Selected Items :" + selected);
                foreach (var item in selectedValues)
                {
                    Log("Item Id:" + item);
                    int avgsalecnsumdays = 0;
                    int reorderdangerdays = 0;
                    int safetystockdays = 0;
                    decimal consumeval = 0;
                    decimal reorderval = 0;
                    decimal safetystockval = 0;
                    decimal Astockval = 0;

                    string qgetdays = $@"select AvgSalesConsumptioninDay,ReOrderDangerLevelinDays,SafetyStockLevelinDays 
                                from muCore_Product  where iMasterId={item}";
                    Log("" + qgetdays);
                    DataSet dsgetdays = ClsDataAcceslayer.GetData1(qgetdays, companyId, ref error);
                    Log("query result count:" + dsgetdays.Tables[0].Rows.Count);
                    if (dsgetdays != null && dsgetdays.Tables[0].Rows.Count > 0)
                    {
                        avgsalecnsumdays = Convert.ToInt32(dsgetdays.Tables[0].Rows[0]["AvgSalesConsumptioninDay"]);
                        reorderdangerdays = Convert.ToInt32(dsgetdays.Tables[0].Rows[0]["ReOrderDangerLevelinDays"]);
                        safetystockdays = Convert.ToInt32(dsgetdays.Tables[0].Rows[0]["SafetyStockLevelinDays"]);//abs
                        Log("For Item :" + item + ", AvgSalesConsumptioninDays :" + avgsalecnsumdays + ",ReOrderDangerLevelinDays :" + reorderdangerdays + ",SafetyStockLevelinDays :" + safetystockdays);
                        if (avgsalecnsumdays == 0)
                        {
                            // return Json(new { status = false, Message = "Please Enter valid Avg Sales/Consumption in Day(s) in Item master" });
                            Log("Please Enter valid Avg Sales/Consumption in Day(s) in Item master for item:" + item);
                        }
                        else
                        {
                            string qgetvalue = $@"select isnull(cast(abs(sum(fQuantity)) as decimal(18,2)),0) fq
                                from tCore_Header_0 h 
                                join tCore_Data_0 d on d.iHeaderId = h.iHeaderId
                                join tCore_Indta_0 ind on ind.iBodyId = d.iBodyId
                                where fQuantity < 0 and  iDate >= dbo.DateToInt(DATEADD(day,-{Convert.ToInt32(dsgetdays.Tables[0].Rows[0]["AvgSalesConsumptioninDay"])}, GETDATE())) 
                                and bUpdateStocks=1
                                and ind.iProduct={item} ";
                            Log("" + qgetvalue);
                            DataSet dsgetvalue = ClsDataAcceslayer.GetData1(qgetvalue, companyId, ref error);
                            Log("Query result count:" + dsgetvalue.Tables[0].Rows.Count);
                            if (dsgetvalue != null && dsgetvalue.Tables[0].Rows.Count > 0)
                            {
                                consumeval = Convert.ToDecimal(dsgetvalue.Tables[0].Rows[0]["fq"]);
                                reorderval = (consumeval / avgsalecnsumdays) * reorderdangerdays;
                                safetystockval = (consumeval / avgsalecnsumdays) * safetystockdays;
                            }
                            string qgetstock = $@"select isnull(sum(fQiss+fQrec),0) stock from vCore_ibals_0 where iProduct={item}";
                            Log("stock query:" + qgetstock);
                            DataSet dsgetstock = ClsDataAcceslayer.GetData1(qgetstock, companyId, ref error);
                            Log("Query result count:" + dsgetstock.Tables[0].Rows.Count);
                            if (dsgetstock != null && dsgetstock.Tables[0].Rows.Count > 0)
                            {
                                Astockval = Convert.ToDecimal(dsgetstock.Tables[0].Rows[0]["stock"]);
                                Log("Actual Stock:" + Astockval);
                            }
                            Log("safetystock Stock:" + safetystockval);
                            if (Astockval < safetystockval)
                            {
                                //string qitemdetails = $@"Select Item,a.iMasterId Itemid,a.UnitId,a.sCode,DefaultBaseUnit Units,isnull(Description,'') Description,
                                //    TankMaster,TankCapacity,ClosingStock,TopUpTankCapacity from
                                //    (select mp.sName Item,mp.sCode,mp.iMasterId,mu.sName DefaultBaseUnit,mu.iMasterId UnitId,sDescription Description,
                                //     iProductType,ProductType from mCore_Product mp
                                //    join muCore_Product_Units mpu on mpu.iMasterId=mp.iMasterId
                                //    join mCore_Units mu on mu.iMasterId=iDefaultBaseUnit
                                //    join muCore_Product mup on mup.iMasterId=mp.iMasterId) a join
                                //    (select w.sName TankMaster,iProduct,isnull(mw.Capacity,0) TankCapacity,sum(fQiss+fQrec) ClosingStock,
                                //    (mw.Capacity-sum(fQiss+fQrec))
                                //    TopUpTankCapacity from tCore_ibals_0  ib join mCore_Warehouse w on ib.iInvTag=w.iMasterId
                                //    join muCore_Warehouse mw on mw.iMasterId=w.iMasterId
                                //    group by w.sName,mw.Capacity,iProduct
                                //    having sum(fQiss+fQrec)>0) b on a.iMasterId=b.iProduct
                                //    where a.iMasterId={item}";

                                string qitemdetails = $@"Select Item,a.Itemid,a.UnitId,a.Code,DefaultBaseUnit Units,
isnull(Description,'') Description,isnull(TankMaster,'') TankMaster,isnull(TankCapacity,0) TankCapacity,isnull(ClosingStock,0) ClosingStock,
isnull(TopUpTankCapacity,0) TopUpTankCapacity,isnull(sVoucherNo,'') sVoucherNo,isnull(v.Whid,0) Whid
--isnull(sProdOrderNo,'') sProdOrderNo,isnull(duedate,'') duedate 
from
(select mp.sName Item,mp.sCode Code,mp.iMasterId Itemid,mu.sName DefaultBaseUnit,mu.iMasterId UnitId,sDescription Description,
iProductType from mCore_Product mp--ProductType
join muCore_Product_Units mpu on mpu.iMasterId=mp.iMasterId
join mCore_Units mu on mu.iMasterId=iDefaultBaseUnit
join muCore_Product mup on mup.iMasterId=mp.iMasterId) a 
left join
(select w.sName TankMaster,iProduct,isnull(mw.Capacity,0) TankCapacity,sum(fQiss+fQrec) ClosingStock,
(mw.Capacity-sum(fQiss+fQrec))
TopUpTankCapacity from vCore_ibals_0  ib join mCore_Warehouse w on ib.iInvTag=w.iMasterId
join muCore_Warehouse mw on mw.iMasterId=w.iMasterId
group by w.sName,mw.Capacity,iProduct
having sum(fQiss+fQrec)>0) b on a.Itemid=b.iProduct

left  join
(
 select top 1 sVoucherNo,iProduct,mw.iMasterId Whid from tCore_Header_0 h join tCore_Data_0 d on h.iHeaderId=d.iHeaderId
join tCore_Indta_0 ind on ind.iBodyId=d.iBodyId
join mCore_Warehouse mw on mw.iMasterId=iInvTag
where iVoucherType=7937
order by len(sVoucherNo) desc,sVoucherNo desc) v on a.Itemid=v.iProduct
where a.Itemid={item}";
                                Log("" + qitemdetails);
                                DataSet dsitemdetails = ClsDataAcceslayer.GetData1(qitemdetails, companyId, ref error);
                                Log("query result count:" + dsitemdetails.Tables[0].Rows.Count);
                                if (dsitemdetails != null && dsitemdetails.Tables[0].Rows.Count > 0)
                                {
                                    for (int i = 0; i < dsitemdetails.Tables[0].Rows.Count; i++)
                                    {
                                        ItemShowModel itemshowmodel = new ItemShowModel();
                                        itemshowmodel.Item = dsitemdetails.Tables[0].Rows[i]["Item"].ToString();
                                        itemshowmodel.Units = dsitemdetails.Tables[0].Rows[i]["Units"].ToString();
                                        itemshowmodel.Description = dsitemdetails.Tables[0].Rows[i]["Description"].ToString();
                                        itemshowmodel.AvailableStockQty = Astockval;
                                        itemshowmodel.SafetyLevelQty = safetystockval;
                                        itemshowmodel.Difference = Astockval - safetystockval;
                                        itemshowmodel.QtytobeProduced = -(Astockval - safetystockval);
                                        itemshowmodel.TankMaster = dsitemdetails.Tables[0].Rows[i]["TankMaster"].ToString();
                                        itemshowmodel.TankCapacity = Convert.ToDecimal(dsitemdetails.Tables[0].Rows[i]["TankCapacity"]);
                                        itemshowmodel.ClosingStock = Convert.ToDecimal(dsitemdetails.Tables[0].Rows[i]["ClosingStock"]);
                                        itemshowmodel.TopUpTankCapacity = Convert.ToDecimal(dsitemdetails.Tables[0].Rows[i]["TopUpTankCapacity"]);
                                        if (dsitemdetails.Tables[0].Rows[0]["sVoucherNo"].ToString() != "")
                                            itemshowmodel.ProductionPlanningStatus = "Raised";
                                        else
                                            itemshowmodel.ProductionPlanningStatus = "Not Raised";
                                        itemshowmodel.ProductionPlanningDocNo = dsitemdetails.Tables[0].Rows[0]["sVoucherNo"].ToString();
                                        //itemshowmodel.DueDate = dsitemdetails.Tables[0].Rows[0]["duedate"].ToString();
                                        itemshowmodel.ItemCode = dsitemdetails.Tables[0].Rows[0]["Code"].ToString();
                                        //itemshowmodel.Remarks = "";                                   
                                        itemshowmodel.Itemid = Convert.ToInt32(dsitemdetails.Tables[0].Rows[i]["Itemid"]);
                                        itemshowmodel.UnitId = Convert.ToInt32(dsitemdetails.Tables[0].Rows[i]["UnitId"]);
                                        itemshowmodel.AvailableStockQty1 = Astockval;
                                        itemshowmodel.SafetyLevelQty1 = safetystockval;
                                        itemshowmodel.Difference1 = Astockval - safetystockval;
                                        itemshowmodel.Whid = Convert.ToInt32(dsitemdetails.Tables[0].Rows[i]["Whid"]);
                                        itemshowmodellist1.Add(itemshowmodel);
                                    }
                                }
                            }
                        }
                    }
                }
                var itemshowmodellist = itemshowmodellist1.OrderBy(_ => _.QtytobeProduced);
                var ddlist = GetWarehouse(companyId);
                return Json(new { status = true, itemshowmodellist, ddlist });
            }
            catch (Exception x)
            {
                return Json(new { status = false, Message = x.Message });
            }
        }
        //GetRMBaseAdditives
        public ActionResult GetRMBaseAdditives(string selected, int companyId, string dt)//, string strwh, int selectedwh , int GroupId
        {
            try
            {
                Log("GetRMBaseAdditives Method");
              //  Log("Selected Group:" + GroupId + ",companyId:" + companyId);
               // string commaseperateditems = "";
                List<ItemShowModel> itemshowmodellist1 = new List<ItemShowModel>();
                List<int> selectedValues = selected.Split(',')
                                  .Select(int.Parse)
                                  .ToList();
                Log("Selected Items :" + selected);
                foreach (var item in selectedValues)
                {
                    Log("Item Id:" + item);
                    int avgsalecnsumdays = 0;
                    int reorderdangerdays = 0;
                    int safetystockdays = 0;
                    decimal consumeval = 0;
                    decimal reorderval = 0;
                    decimal safetystockval = 0;
                    decimal Astockval = 0;

                    string qgetdays = $@"select AvgSalesConsumptioninDay,ReOrderDangerLevelinDays,SafetyStockLevelinDays 
                                from muCore_Product  where iMasterId={item}";
                    Log("" + qgetdays);
                    DataSet dsgetdays = ClsDataAcceslayer.GetData1(qgetdays, companyId, ref error);
                    Log("query result count:" + dsgetdays.Tables[0].Rows.Count);
                    if (dsgetdays != null && dsgetdays.Tables[0].Rows.Count > 0)
                    {
                        avgsalecnsumdays = Convert.ToInt32(dsgetdays.Tables[0].Rows[0]["AvgSalesConsumptioninDay"]);
                        reorderdangerdays = Convert.ToInt32(dsgetdays.Tables[0].Rows[0]["ReOrderDangerLevelinDays"]);
                        safetystockdays = Convert.ToInt32(dsgetdays.Tables[0].Rows[0]["SafetyStockLevelinDays"]);//abs
                        Log("For Item :" + item + ", AvgSalesConsumptioninDays :" + avgsalecnsumdays + ",ReOrderDangerLevelinDays :" + reorderdangerdays + ",SafetyStockLevelinDays :" + safetystockdays);
                        if (avgsalecnsumdays == 0)
                        {
                            //return Json(new { status = false, Message = "Please Enter valid Avg Sales/Consumption in Day(s) in Item master" });
                            Log("Please Enter valid Avg Sales/Consumption in Day(s) in Item master for item:" + item);
                        }
                        else
                        {
                            string qgetvalue = $@"select isnull(cast(abs(sum(fQuantity)) as decimal(18,2)),0) fq
                                from tCore_Header_0 h 
                                join tCore_Data_0 d on d.iHeaderId = h.iHeaderId
                                join tCore_Indta_0 ind on ind.iBodyId = d.iBodyId
                                where fQuantity < 0 and  iDate >= dbo.DateToInt(DATEADD(day,-{Convert.ToInt32(dsgetdays.Tables[0].Rows[0]["AvgSalesConsumptioninDay"])}, GETDATE())) 
                                and bUpdateStocks=1
                                and ind.iProduct={item} ";
                            Log("" + qgetvalue);
                            DataSet dsgetvalue = ClsDataAcceslayer.GetData1(qgetvalue, companyId, ref error);
                            Log("Query result count:" + dsgetvalue.Tables[0].Rows.Count);
                            if (dsgetvalue != null && dsgetvalue.Tables[0].Rows.Count > 0)
                            {
                                consumeval = Convert.ToDecimal(dsgetvalue.Tables[0].Rows[0]["fq"]);
                                reorderval = (consumeval / avgsalecnsumdays) * reorderdangerdays;
                                safetystockval = (consumeval / avgsalecnsumdays) * safetystockdays;
                            }
                            string qgetstock = $@"select isnull(sum(fQiss+fQrec),0) stock from vCore_ibals_0 where iProduct={item}";
                            Log("stock query:" + qgetstock);
                            DataSet dsgetstock = ClsDataAcceslayer.GetData1(qgetstock, companyId, ref error);
                            Log("Query result count:" + dsgetstock.Tables[0].Rows.Count);
                            if (dsgetstock != null && dsgetstock.Tables[0].Rows.Count > 0)
                            {
                                Astockval = Convert.ToDecimal(dsgetstock.Tables[0].Rows[0]["stock"]);
                                Log("Actual Stock:" + Astockval);
                            }
                            Log("safetystock Stock:" + safetystockval);
                            if (Astockval < safetystockval)
                            {
                                //string qitemdetails = $@"select mp.sName Item,mp.iMasterId Itemid,mp.sCode,mu.sName DefaultBaseUnit,
                                //    mu.iMasterId UnitId,sDescription Description 
                                //  from mCore_Product mp
                                //  join muCore_Product_Units mpu on mpu.iMasterId=mp.iMasterId
                                //  join mCore_Units mu on mu.iMasterId=iDefaultBaseUnit
                                //  join muCore_Product mup on mup.iMasterId=mp.iMasterId
                                //  where mp.iMasterId={item}";
                                string qitemdetails = $@"select Item,Itemid,sCode,UnitId,DefaultBaseUnit,Description,isnull(sVoucherNo,'') 
sVoucherNo,isnull(TankMaster,'') TankMaster,isnull(TankCapacity,0) TankCapacity,isnull(ClosingStock,0) ClosingStock,
isnull(TopUpTankCapacity,0) TopUpTankCapacity,isnull(l.Whid,0) Whid from
(select mp.sName Item,mp.iMasterId Itemid,mp.sCode,mu.sName DefaultBaseUnit,
mu.iMasterId UnitId,sDescription Description 
from mCore_Product mp
join muCore_Product_Units mpu on mpu.iMasterId=mp.iMasterId
join mCore_Units mu on mu.iMasterId=iDefaultBaseUnit
join muCore_Product mup on mup.iMasterId=mp.iMasterId
) k 
left join
(select w.sName TankMaster,iProduct,isnull(mw.Capacity,0) TankCapacity,sum(fQiss+fQrec) ClosingStock,
(mw.Capacity-sum(fQiss+fQrec))
TopUpTankCapacity from vCore_ibals_0  ib join mCore_Warehouse w on ib.iInvTag=w.iMasterId
join muCore_Warehouse mw on mw.iMasterId=w.iMasterId
group by w.sName,mw.Capacity,iProduct
having sum(fQiss+fQrec)>0) b on k.Itemid=b.iProduct
	left  join
( select top 1 sVoucherNo,iProduct,mw.iMasterId Whid from tCore_Header_0 h join tCore_Data_0 d on h.iHeaderId=d.iHeaderId
join tCore_Indta_0 ind on ind.iBodyId=d.iBodyId
join mCore_Warehouse mw on mw.iMasterId=iInvTag
where iVoucherType=7937
order by len(sVoucherNo) desc,sVoucherNo desc) l on k.Itemid=l.iProduct
where Itemid={item}";
                                Log("" + qitemdetails);
                                DataSet dsitemdetails = ClsDataAcceslayer.GetData1(qitemdetails, companyId, ref error);
                                Log("query result count:" + dsitemdetails.Tables[0].Rows.Count);
                                if (dsitemdetails != null && dsitemdetails.Tables[0].Rows.Count > 0)
                                {
                                    for (int i = 0; i < dsitemdetails.Tables[0].Rows.Count; i++)
                                    {
                                        ItemShowModel itemshowmodel = new ItemShowModel();
                                        itemshowmodel.Item = dsitemdetails.Tables[0].Rows[i]["Item"].ToString();
                                        //itemshowmodel.Warehouse = strwh;
                                        itemshowmodel.Itemid = Convert.ToInt32(dsitemdetails.Tables[0].Rows[i]["Itemid"]);
                                        itemshowmodel.UnitId = Convert.ToInt32(dsitemdetails.Tables[0].Rows[i]["UnitId"]);
                                        itemshowmodel.Units = dsitemdetails.Tables[0].Rows[i]["DefaultBaseUnit"].ToString();
                                        itemshowmodel.Description = dsitemdetails.Tables[0].Rows[i]["Description"].ToString();
                                        itemshowmodel.AvailableStockQty = Astockval;
                                        itemshowmodel.SafetyLevelQty = safetystockval;
                                        itemshowmodel.Difference = Astockval - safetystockval;
                                        itemshowmodel.QtytobeProduced = -(Astockval - safetystockval);
                                        if (dsitemdetails.Tables[0].Rows[i]["sVoucherNo"].ToString() != "")
                                            itemshowmodel.ProductionPlanningStatus = "Raised";
                                        else
                                            itemshowmodel.ProductionPlanningStatus = "Not Raised";
                                        itemshowmodel.ProductionPlanningDocNo = dsitemdetails.Tables[0].Rows[i]["sVoucherNo"].ToString();
                                        //itemshowmodel.DueDate = dsitemdetails.Tables[0].Rows[0]["duedate"].ToString();
                                        itemshowmodel.ItemCode = dsitemdetails.Tables[0].Rows[i]["sCode"].ToString();
                                        itemshowmodel.Whid = Convert.ToInt32(dsitemdetails.Tables[0].Rows[i]["Whid"]);
                                        itemshowmodellist1.Add(itemshowmodel);
                                    }
                                }
                            }
                        }
                    }
                }
                var itemshowmodellist = itemshowmodellist1.OrderBy(_ => _.QtytobeProduced);
                var ddlist = GetWarehouse(companyId);
                return Json(new { status = true, itemshowmodellist, ddlist });
            }
            catch (Exception X)
            {
                return Json(new { status = false, Message = X.Message });
            }
        }
        //GetPm
        public ActionResult GetPm(string selected, int companyId, string dt)//, string strwh, int selectedwh
        {
            try
            {
                Log("GetPm Method");
             //   Log("Selected Group:" + GroupId + ",companyId:" + companyId);
               // string commaseperateditems = "";
                List<ItemShowModel> itemshowmodellist1 = new List<ItemShowModel>();
                List<int> selectedValues = selected.Split(',')
                                  .Select(int.Parse)
                                  .ToList();
                Log("Selected Items :" + selected);
                foreach (var item in selectedValues)
                {
                    Log("Item Id:" + item);
                    int avgsalecnsumdays = 0;
                    int reorderdangerdays = 0;
                    int safetystockdays = 0;
                    decimal consumeval = 0;
                    decimal reorderval = 0;
                    decimal safetystockval = 0;
                    decimal Astockval = 0;

                    string qgetdays = $@"select AvgSalesConsumptioninDay,ReOrderDangerLevelinDays,SafetyStockLevelinDays 
                                from muCore_Product  where iMasterId={item}";
                    Log("" + qgetdays);
                    DataSet dsgetdays = ClsDataAcceslayer.GetData1(qgetdays, companyId, ref error);
                    Log("query result count:" + dsgetdays.Tables[0].Rows.Count);
                    if (dsgetdays != null && dsgetdays.Tables[0].Rows.Count > 0)
                    {
                        avgsalecnsumdays = Convert.ToInt32(dsgetdays.Tables[0].Rows[0]["AvgSalesConsumptioninDay"]);
                        reorderdangerdays = Convert.ToInt32(dsgetdays.Tables[0].Rows[0]["ReOrderDangerLevelinDays"]);
                        safetystockdays = Convert.ToInt32(dsgetdays.Tables[0].Rows[0]["SafetyStockLevelinDays"]);//abs
                        Log("For Item :" + item + ", AvgSalesConsumptioninDays :" + avgsalecnsumdays + ",ReOrderDangerLevelinDays :" + reorderdangerdays + ",SafetyStockLevelinDays :" + safetystockdays);
                        if (avgsalecnsumdays == 0)
                        {// return Json(new { status = false, Message = "Please Enter valid Avg Sales/Consumption in Day(s) in Item master" });
                            Log("Please Enter valid Avg Sales/Consumption in Day(s) in Item master for item:" + item);
                        }
                        else
                        {
                            string qgetvalue = $@"select isnull(cast(abs(sum(fQuantity)) as decimal(18,2)),0) fq
                                from tCore_Header_0 h 
                                join tCore_Data_0 d on d.iHeaderId = h.iHeaderId
                                join tCore_Indta_0 ind on ind.iBodyId = d.iBodyId
                                where fQuantity < 0 and  iDate >= dbo.DateToInt(DATEADD(day,-{Convert.ToInt32(dsgetdays.Tables[0].Rows[0]["AvgSalesConsumptioninDay"])}, GETDATE())) 
                                and bUpdateStocks=1
                                and ind.iProduct={item}";
                            Log("" + qgetvalue);
                            DataSet dsgetvalue = ClsDataAcceslayer.GetData1(qgetvalue, companyId, ref error);
                            Log("Query result count:" + dsgetvalue.Tables[0].Rows.Count);
                            if (dsgetvalue != null && dsgetvalue.Tables[0].Rows.Count > 0)
                            {
                                consumeval = Convert.ToDecimal(dsgetvalue.Tables[0].Rows[0]["fq"]);
                                reorderval = (consumeval / avgsalecnsumdays) * reorderdangerdays;
                                safetystockval = (consumeval / avgsalecnsumdays) * safetystockdays;
                            }
                            string qgetstock = $@"select isnull(sum(fQiss+fQrec),0) stock from vCore_ibals_0 where iProduct={item}";
                            Log("stock query:" + qgetstock);
                            DataSet dsgetstock = ClsDataAcceslayer.GetData1(qgetstock, companyId, ref error);
                            Log("Query result count:" + dsgetstock.Tables[0].Rows.Count);
                            if (dsgetstock != null && dsgetstock.Tables[0].Rows.Count > 0)
                            {
                                Astockval = Convert.ToDecimal(dsgetstock.Tables[0].Rows[0]["stock"]);
                                Log("Actual Stock:" + Astockval);
                            }
                            Log("safetystock Stock:" + safetystockval);
                            if (Astockval < safetystockval)
                            {
                                //string qitemdetails = $@"select mp.sName Item,mp.iMasterId Itemid,mp.sCode,mu.sName DefaultBaseUnit,
                                //    mu.iMasterId UnitId,sDescription Description 
                                //  from mCore_Product mp
                                //  join muCore_Product_Units mpu on mpu.iMasterId=mp.iMasterId
                                //  join mCore_Units mu on mu.iMasterId=iDefaultBaseUnit
                                //  join muCore_Product mup on mup.iMasterId=mp.iMasterId
                                //  where mp.iMasterId={item}";
                                string qitemdetails = $@"select Item,Itemid,sCode,UnitId,DefaultBaseUnit,Description,isnull(sVoucherNo,'') 
sVoucherNo,isnull(TankMaster,'') TankMaster,isnull(TankCapacity,0) TankCapacity,isnull(ClosingStock,0) ClosingStock,
isnull(TopUpTankCapacity,0) TopUpTankCapacity,isnull(l.Whid,0) Whid from
(select mp.sName Item,mp.iMasterId Itemid,mp.sCode,mu.sName DefaultBaseUnit,
mu.iMasterId UnitId,sDescription Description 
from mCore_Product mp
join muCore_Product_Units mpu on mpu.iMasterId=mp.iMasterId
join mCore_Units mu on mu.iMasterId=iDefaultBaseUnit
join muCore_Product mup on mup.iMasterId=mp.iMasterId
) k
left join
(select w.sName TankMaster,iProduct,isnull(mw.Capacity,0) TankCapacity,sum(fQiss+fQrec) ClosingStock,
(mw.Capacity-sum(fQiss+fQrec))
TopUpTankCapacity from vCore_ibals_0  ib join mCore_Warehouse w on ib.iInvTag=w.iMasterId
join muCore_Warehouse mw on mw.iMasterId=w.iMasterId
group by w.sName,mw.Capacity,iProduct
having sum(fQiss+fQrec)>0) b on k.Itemid=b.iProduct
	left  join
(select top 1 sVoucherNo,iProduct,mw.iMasterId Whid from tCore_Header_0 h join tCore_Data_0 d on h.iHeaderId=d.iHeaderId
join tCore_Indta_0 ind on ind.iBodyId=d.iBodyId
join mCore_Warehouse mw on mw.iMasterId=iInvTag
where iVoucherType=7937
order by len(sVoucherNo) desc,sVoucherNo desc) l on k.Itemid=l.iProduct
where Itemid={item}";
                                Log("" + qitemdetails);
                                DataSet dsitemdetails = ClsDataAcceslayer.GetData1(qitemdetails, companyId, ref error);
                                Log("query result count:" + dsitemdetails.Tables[0].Rows.Count);
                                if (dsitemdetails != null && dsitemdetails.Tables[0].Rows.Count > 0)
                                {
                                    for (int i = 0; i < dsitemdetails.Tables[0].Rows.Count; i++)
                                    {
                                        ItemShowModel itemshowmodel = new ItemShowModel();
                                        itemshowmodel.Item = dsitemdetails.Tables[0].Rows[i]["Item"].ToString();
                                        //itemshowmodel.Warehouse = strwh;
                                        itemshowmodel.Itemid = Convert.ToInt32(dsitemdetails.Tables[0].Rows[i]["Itemid"]);
                                        itemshowmodel.UnitId = Convert.ToInt32(dsitemdetails.Tables[0].Rows[i]["UnitId"]);
                                        itemshowmodel.Units = dsitemdetails.Tables[0].Rows[i]["DefaultBaseUnit"].ToString();
                                        itemshowmodel.Description = dsitemdetails.Tables[0].Rows[i]["Description"].ToString();
                                        itemshowmodel.AvailableStockQty = Astockval;
                                        itemshowmodel.SafetyLevelQty = safetystockval;
                                        itemshowmodel.Difference = Astockval - safetystockval;
                                        itemshowmodel.QtytobeProduced = -(Astockval - safetystockval);
                                        if (dsitemdetails.Tables[0].Rows[i]["sVoucherNo"].ToString() != "")
                                            itemshowmodel.ProductionPlanningStatus = "Raised";
                                        else
                                            itemshowmodel.ProductionPlanningStatus = "Not Raised";
                                        itemshowmodel.ProductionPlanningDocNo = dsitemdetails.Tables[0].Rows[i]["sVoucherNo"].ToString();
                                        //itemshowmodel.DueDate = dsitemdetails.Tables[0].Rows[0]["duedate"].ToString();
                                        itemshowmodel.ItemCode = dsitemdetails.Tables[0].Rows[i]["sCode"].ToString();
                                        itemshowmodel.Whid = Convert.ToInt32(dsitemdetails.Tables[0].Rows[i]["Whid"]);
                                        itemshowmodellist1.Add(itemshowmodel);
                                    }
                                }
                            }
                        }
                    }
                }
                var itemshowmodellist = itemshowmodellist1.OrderBy(_ => _.QtytobeProduced);
                var ddlist = GetWarehouse(companyId);
                return Json(new { status = true, itemshowmodellist, ddlist });
            }
            catch (Exception X)
            {
                return Json(new { status = false, Message = X.Message });
            }
        }
        //SFGPost
        public ActionResult SFGPost(List<ItemShowModel> collectedData, string SessionId, int selectedincharge, int cid)
        {
            try
            {
                #region Header
                string Message = "";
                Hashtable header = new Hashtable();
                header = new Hashtable
                        {
                           // { "DocNo", VoucherExistsVoucher},
                           // { "Date", CIssue.Date },
                            { "CustomerAC__Id", selectedincharge },
                            { "Branch__Id", 4}


                        };
                #endregion
                #region Body     
                List<Hashtable> body = new List<Hashtable>();
                Hashtable row1 = new Hashtable();
                foreach (var item in collectedData)
                {
                    int Idate = ClsDataAcceslayer.GetDateToInt(Convert.ToDateTime(item.DueDate));
                    row1 = new Hashtable
                                      {
                                           {"Item__Id", item.Itemid },
                                           {"Unit__Id",item.UnitId  },           //item.Units                                
                                           {"Quantity", item.QtytobeProduced },
                                           {"AvailableStock", item.AvailableStockQty1},
                                           {"SafetyLevelQty", item.SafetyLevelQty1 },
                                           {"Difference", item.Difference1 },
                                           {"QuantitytobeProduced", item.QtytobeProduced },
                                           {"TankTopUpCapacity", item.TopUpTankCapacity },
                                           {"DueDate", Idate }
                                          // {"SafetyLevelQty", item.SafetyLevelQty },
                    };
                    body.Add(row1);
                }
                #endregion
                var postingData = new PostingData();
                postingData.data.Add(new Hashtable { { "Header", header }, { "Body", body } });
                string sContent = JsonConvert.SerializeObject(postingData);
                Log("Posting Updating sContent" + sContent);
                Log("Before Response");
                Log("Api URL:" + ProdPlan);
                var response = Focus8API.Post(ProdPlan, sContent, SessionId, ref error);
                Log("response:" + response);
                Log("After Response");
                if (response != null)
                {
                    var responseData = JsonConvert.DeserializeObject<APIResponse.PostResponse>(response);
                    if (responseData.result == -1)
                    {
                        Message = $"Posting Failed : {responseData.message } \n";
                        Log("Error posting failed:" + Message);
                    }
                    else
                    {
                        var docNo = Convert.ToString(responseData.data[0]["VoucherNo"]);
                        Message = "Posted into Production Planning successfully";
                    }
                }
                return Json(new { status = true, Message });
            }
            catch (Exception x)
            {
                return Json(new { status = false, Message = x.Message });
            }
        }

        public ActionResult FGPost(List<ItemShowModel> collectedData, string SessionId, int selectedincharge, int cid)
        {
            try
            {
                string Message = "";
                Log("FG Post Method");
                #region Header
                //string Message = "";
                Hashtable header = new Hashtable();
                header = new Hashtable
                        {
                           // { "DocNo", VoucherExistsVoucher},
                           // { "Date", CIssue.Date },
                            { "CustomerAC__Id",selectedincharge },
                            { "Branch__Id", 4}

                        };
                #endregion
                #region Body     
                List<Hashtable> body = new List<Hashtable>();
                Hashtable row1 = new Hashtable();
                int cc = 0;
                int kk = 0;
                foreach (var item in collectedData)
                {
                    int dd = 0;
                    int ibom = 0;
                    string qbommap = $@"select r.iBOM,p.sName from muCore_Product_Replenishment r join 
									mCore_Product p on r.iMasterId=p.iMasterId
									where p.iMasterId={item.Itemid}";
                    DataSet dsbommap = ClsDataAcceslayer.GetData1(qbommap, cid, ref error);
                    if (dsbommap != null && dsbommap.Tables[0].Rows.Count > 0)
                    {
                        ibom = Convert.ToInt32(dsbommap.Tables[0].Rows[0]["iBOM"]);
                    }
                    if (ibom == 0)
                    {
                        cc++;
                        Message += $"{dsbommap.Tables[0].Rows[0]["iBOM"].ToString()} not mapped to BOM.\n";
                        Log($"'{dsbommap.Tables[0].Rows[0]["iBOM"].ToString()}' not mapped to BOM");
                    }
                    else
                    {
                        int Idate = ClsDataAcceslayer.GetDateToInt(Convert.ToDateTime(item.DueDate));
                        string qgetbomip = $@"select iBOM,iVersion,bh.sName BomName,mp.iMasterId ItemId,mp.sName Item,mp.iProductType,bb.fQty,bInput
                                    from muCore_Product_Replenishment mp1--bh.sName output
									join mMRP_BomVariantHeader vh on vh.iVariantId=iBOM
									join mMRP_BomHeader bh on bh.iBomId=vh.iBomId and bh.iStatus<>5
									join mMRP_BOMBody bb on bb.iVariantId=vh.iVariantId 
									join mCore_Product mp on mp.iMasterId=bb.iProductId
									where mp1.iMasterId={item.Itemid} order by bInput--and bInput=1";
                        DataSet dsgetbomip = ClsDataAcceslayer.GetData1(qgetbomip, cid, ref error);
                        List<BomModel> bommodellist = new List<BomModel>();
                        if (dsgetbomip != null && dsgetbomip.Tables[0].Rows.Count > 0)
                        {
                            for (int i = 0; i < dsgetbomip.Tables[0].Rows.Count; i++)
                            {
                                BomModel bommodel = new BomModel();
                                bommodel.iBOM = Convert.ToInt32(dsgetbomip.Tables[0].Rows[i]["iBOM"]);
                                bommodel.iVersion = Convert.ToInt32(dsgetbomip.Tables[0].Rows[i]["iVersion"]);
                                bommodel.BomName = dsgetbomip.Tables[0].Rows[i]["BomName"].ToString();
                                bommodel.ItemId = Convert.ToInt32(dsgetbomip.Tables[0].Rows[i]["ItemId"]);
                                bommodel.Item = dsgetbomip.Tables[0].Rows[i]["Item"].ToString();
                                bommodel.iProductType = Convert.ToInt32(dsgetbomip.Tables[0].Rows[i]["iProductType"]);
                                bommodel.fQty = Convert.ToDecimal(dsgetbomip.Tables[0].Rows[i]["fQty"]);
                                bommodel.bInput = Convert.ToInt32(dsgetbomip.Tables[0].Rows[i]["bInput"]);
                                bommodellist.Add(bommodel);
                            }
                        }
                        var bomoplst = bommodellist.Where(_ => _.bInput == 0).ToList();
                        var FGitemquantity = bomoplst.FirstOrDefault().fQty;
                        var bomipsfglst = bommodellist.Where(_ => (_.bInput == 1 && _.iProductType == 3)).ToList();
                        var bomipotherlst = bommodellist.Where(_ => (_.bInput == 1 && _.iProductType != 3)).ToList();
                        //check for SFG
                        Log($"BOM name :{bomoplst.FirstOrDefault().BomName}({bomoplst.FirstOrDefault().iBOM},Version-{bomoplst.FirstOrDefault().iVersion}) ,FG item name:{bomoplst.FirstOrDefault().Item}({bomoplst.FirstOrDefault().ItemId}) ,quantity :{bomoplst.FirstOrDefault().fQty}");
                        if (bomipsfglst.Count() > 0)
                        {
                            foreach (var ipitem in bomipsfglst)
                            {
                                //if (ipitem.iProductType == 3)
                                //{
                                Log($"BOM name :{ipitem.BomName}({ipitem.iBOM},Version-{ipitem.iVersion}) ,SFG item name:{ipitem.Item}({ipitem.ItemId}) ,quantity :{ipitem.fQty}");
                                var SFGitemquantity = ipitem.fQty;
                                var singleitemqty = SFGitemquantity / FGitemquantity;
                                var reqqty = item.QtytobeProduced * singleitemqty;
                                Log($"Required quantity for SFG item {ipitem.Item} :{ reqqty}");
                                string qcrntstk = $@"select case when sum(fQiss+fQrec)  is null then 0
									when sum(fQiss+fQrec)<0 then 0 else sum(fQiss+fQrec) end Closingstock from vCore_ibals_0 where iProduct=
                                    {ipitem.ItemId}";
                                DataSet dscrntstk = ClsDataAcceslayer.GetData1(qcrntstk, cid, ref error);
                                Log($"Available stock for SFG item {ipitem.Item}: {Convert.ToDecimal(dscrntstk.Tables[0].Rows[0]["Closingstock"])}");
                                if (Convert.ToDecimal(dscrntstk.Tables[0].Rows[0]["Closingstock"]) == 0 || Convert.ToDecimal(dscrntstk.Tables[0].Rows[0]["Closingstock"]) < reqqty)
                                {
                                    dd++;
                                    kk++;
                                    Message += $"SFG item '{ipitem.Item}' has low stock balance.\n";
                                    Log($"SFG item '{ipitem.Item}' has low stock balance");
                                }
                                //}
                            }
                        }
                        //else
                        //{
                        //    row1 = new Hashtable
                        //      {
                        //           {"Item__Id", item.Itemid },
                        //           {"Unit__Id",item.UnitId  },           //item.Units                                
                        //           {"Quantity", item.QtytobeProduced },
                        //           {"AvailableStock", item.AvailableStockQty},
                        //           {"SafetyLevelQty", item.SafetyLevelQty },
                        //           {"Difference", item.Difference },
                        //           {"QuantitytobeProduced", item.QtytobeProduced },
                        //           //{"TankTopUpCapacity", item.TopUpTankCapacity },
                        //           {"DueDate", Idate }
                        //          // {"SafetyLevelQty", item.SafetyLevelQty },
                        //      };
                        //    body.Add(row1);
                        //}

                        //else
                        //{
                        //    row1 = new Hashtable
                        //          {
                        //               {"Item__Id", item.Itemid },
                        //               {"Unit__Id",item.UnitId  },           //item.Units                                
                        //               {"Quantity", item.QtytobeProduced },
                        //               {"AvailableStock", item.AvailableStockQty},
                        //               {"SafetyLevelQty", item.SafetyLevelQty },
                        //               {"Difference", item.Difference },
                        //               {"QuantitytobeProduced", item.QtytobeProduced },
                        //               //{"TankTopUpCapacity", item.TopUpTankCapacity },
                        //               {"DueDate", Idate }
                        //              // {"SafetyLevelQty", item.SafetyLevelQty },
                        //         };
                        //    body.Add(row1);
                        //}

                        //                       if (Convert.ToInt32(dsgetbomip.Tables[0].Rows[i]["iProductType"]) == 3)
                        //                       {

                        //                           var qty = item.QtytobeProduced * Convert.ToDecimal(dsgetbomip.Tables[0].Rows[i]["fQty"]);
                        //                           Log("Quantity to be required :" + item.QtytobeProduced + "*" + Convert.ToDecimal(dsgetbomip.Tables[0].Rows[i]["fQty"]) + "=" + qty);
                        //                           string qcrntstk = $@"select case when sum(fQiss+fQrec)  is null then 0
                        //when sum(fQiss+fQrec)<0 then 0 else sum(fQiss+fQrec) end Closingstock from vCore_ibals_0 where iProduct=
                        //                           {Convert.ToInt32(dsgetbomip.Tables[0].Rows[i]["Input"])}";
                        //                           DataSet dscrntstk = ClsDataAcceslayer.GetData1(qcrntstk, cid, ref error);
                        //                           Log("Available stock:" + Convert.ToDecimal(dscrntstk.Tables[0].Rows[0]["Closingstock"]));
                        //                           if (Convert.ToDecimal(dscrntstk.Tables[0].Rows[0]["Closingstock"]) == 0 || Convert.ToDecimal(dscrntstk.Tables[0].Rows[0]["Closingstock"]) < qty)
                        //                           {
                        //                               Log($"SFG '{dsgetbomip.Tables[0].Rows[i]["Item"].ToString()}' dont have enough stock to raise ,so no posting done for it ");
                        //                           }
                        if (dd == 0)
                        {
                            row1 = new Hashtable
                                      {
                                           {"Item__Id", item.Itemid },
                                           {"Unit__Id",item.UnitId  },           //item.Units                                
                                           {"Quantity", item.QtytobeProduced },
                                           {"AvailableStock", item.AvailableStockQty},
                                           {"SafetyLevelQty", item.SafetyLevelQty },
                                           {"Difference", item.Difference },
                                           {"QuantitytobeProduced", item.QtytobeProduced },
                                           //{"TankTopUpCapacity", item.TopUpTankCapacity },
                                           {"DueDate", Idate }
                                          // {"SafetyLevelQty", item.SafetyLevelQty },
                                      };
                            body.Add(row1);
                        }

                        //                       }
                        //                       //for others
                        //                       else
                        //                       {
                        //                           row1 = new Hashtable
                        //                             {
                        //                                  {"Item__Id", item.Itemid },
                        //                                  {"Unit__Id",item.UnitId  },           //item.Units                                
                        //                                  {"Quantity", item.QtytobeProduced },
                        //                                  {"AvailableStock", item.AvailableStockQty},
                        //                                  {"SafetyLevelQty", item.SafetyLevelQty },
                        //                                  {"Difference", item.Difference },
                        //                                  {"QuantitytobeProduced", item.QtytobeProduced },
                        //                                  //{"TankTopUpCapacity", item.TopUpTankCapacity },
                        //                                  {"DueDate", Idate }
                        //                                 // {"SafetyLevelQty", item.SafetyLevelQty },
                        //                            };
                        //                           body.Add(row1);
                        //                       }
                    }
                }



                // }
                //cc++;
                // }
                #endregion
                if (cc == collectedData.Count())
                {
                    Log("All FG(s) are not mapped with BOM");
                    return Json(new { status = false, Message = "Item(s) are not mapped with BOM" });
                }
                if (body.Count() == 0)
                {
                    Log("All SFG(s) are not having suffient balance");
                    return Json(new { status = false, Message = "SFG Item(s) don't have sufficient stock" });
                }

                var postingData = new PostingData();
                postingData.data.Add(new Hashtable { { "Header", header }, { "Body", body } });
                string sContent = JsonConvert.SerializeObject(postingData);
                Log("Posting Updating sContent" + sContent);
                Log("Before Response");
                Log("Api URL:" + ProdPlan);
                var response = Focus8API.Post(ProdPlan, sContent, SessionId, ref error);
                Log("response:" + response);
                Log("After Response");
                if (response != null)
                {
                    var responseData = JsonConvert.DeserializeObject<APIResponse.PostResponse>(response);
                    if (responseData.result == -1)
                    {
                        Message += $"Posting Failed : {responseData.message } \n";
                        Log("Error posting failed:" + Message);
                    }
                    else
                    {
                        var docNo = Convert.ToString(responseData.data[0]["VoucherNo"]);
                        if (cc > 0 || kk > 0) { Message += "Partially Posted into Production Planning successfully.\n"; }
                        if (cc == 0 && kk == 0) { Message += "Posted into Production Planning successfully.\n"; }
                    }
                }
                return Json(new { status = true, Message });

            }
            catch (Exception x)
            {
                return Json(new { status = false, Message = x.Message });
            }
        }
        //RMBaseOilsPost
        public ActionResult RMBaseOilsPost(List<ItemShowModel> collectedData, string SessionId, int cid)
        {
            try
            {
                Log("RM BaseOils Post Method");
                #region Header
                string Message = "";
                Hashtable header = new Hashtable();
                header = new Hashtable
                        {
                           // { "DocNo", VoucherExistsVoucher},
                           // { "Date", CIssue.Date },
                            //{ "CustomerAC__Id", selectedincharge },
                            { "Branch__Id", 4}


                        };
                #endregion
                #region body
                List<Hashtable> body = new List<Hashtable>();
                Hashtable row1 = new Hashtable();
                foreach (var item in collectedData)
                {
                    // int Idate = ClsDataAcceslayer.GetDateToInt(Convert.ToDateTime(item.DueDate));
                    row1 = new Hashtable
                                      {
                                           {"Warehouse__Id",item.Whid},
                                           {"Item__Id", item.Itemid },
                                           {"Unit__Id",item.UnitId  },           //item.Units                                
                                           {"Quantity", item.QtytobeProduced },
                                           //{"AvailableStock", item.AvailableStockQty},
                                           //{"SafetyLevelQty", item.SafetyLevelQty },
                                           //{"Difference", item.Difference },
                                           //{"QuantitytobeProduced", item.QtytobeProduced },
                                           //{"TankTopUpCapacity", item.TopUpTankCapacity },
                                           //{"DueDate", Idate }
                                          // {"SafetyLevelQty", item.SafetyLevelQty },
                    };
                    body.Add(row1);
                }
                #endregion
                var postingData = new PostingData();
                postingData.data.Add(new Hashtable { { "Header", header }, { "Body", body } });
                string sContent = JsonConvert.SerializeObject(postingData);
                Log("Posting Updating sContent" + sContent);
                Log("Before Response");
                Log("Api URL:" + ProdPlan);
                var response = Focus8API.Post(PurInd, sContent, SessionId, ref error);
                Log("response:" + response);
                Log("After Response");
                if (response != null)
                {
                    var responseData = JsonConvert.DeserializeObject<APIResponse.PostResponse>(response);
                    if (responseData.result == -1)
                    {
                        Message = $"Posting Failed : {responseData.message } \n";
                        Log("Error posting failed:" + Message);
                    }
                    else
                    {
                        var docNo = Convert.ToString(responseData.data[0]["VoucherNo"]);
                        Message = "Posted into Purchase Indent successfully";
                    }
                }
                return Json(new { status = true, Message });

            }
            catch (Exception x)
            {
                return Json(new { status = false, Message = x.Message });
            }
        }
        //RMAdditivesPost
        public ActionResult RMAdditivesPost(List<ItemShowModel> collectedData, string SessionId, int cid)
        {
            try
            {
                Log("RM Additives Post Method");
                #region Header
                string Message = "";
                Hashtable header = new Hashtable();
                header = new Hashtable
                        {
                           // { "DocNo", VoucherExistsVoucher},
                           // { "Date", CIssue.Date },
                            //{ "CustomerAC__Id", selectedincharge },
                            { "Branch__Id", 4}


                        };
                #endregion
                #region body
                List<Hashtable> body = new List<Hashtable>();
                Hashtable row1 = new Hashtable();
                foreach (var item in collectedData)
                {
                    // int Idate = ClsDataAcceslayer.GetDateToInt(Convert.ToDateTime(item.DueDate));
                    row1 = new Hashtable
                                      {
                                           {"Warehouse__Id",item.Whid},
                                           {"Item__Id", item.Itemid },
                                           {"Unit__Id",item.UnitId  },           //item.Units                                
                                           {"Quantity", item.QtytobeProduced },
                                           //{"AvailableStock", item.AvailableStockQty},
                                           //{"SafetyLevelQty", item.SafetyLevelQty },
                                           //{"Difference", item.Difference },
                                           //{"QuantitytobeProduced", item.QtytobeProduced },
                                           //{"TankTopUpCapacity", item.TopUpTankCapacity },
                                           //{"DueDate", Idate }
                                          // {"SafetyLevelQty", item.SafetyLevelQty },
                    };
                    body.Add(row1);
                }
                #endregion
                var postingData = new PostingData();
                postingData.data.Add(new Hashtable { { "Header", header }, { "Body", body } });
                string sContent = JsonConvert.SerializeObject(postingData);
                Log("Posting Updating sContent" + sContent);
                Log("Before Response");
                Log("Api URL:" + ProdPlan);
                var response = Focus8API.Post(PurInd, sContent, SessionId, ref error);
                Log("response:" + response);
                Log("After Response");
                if (response != null)
                {
                    var responseData = JsonConvert.DeserializeObject<APIResponse.PostResponse>(response);
                    if (responseData.result == -1)
                    {
                        Message = $"Posting Failed : {responseData.message } \n";
                        Log("Error posting failed:" + Message);
                    }
                    else
                    {
                        var docNo = Convert.ToString(responseData.data[0]["VoucherNo"]);
                        Message = "Posted into Purchase Indent successfully";
                    }
                }
                return Json(new { status = true, Message });
            }
            catch (Exception x)
            {
                return Json(new { status = false, Message = x.Message });
            }
        }
        //PMPost
        public ActionResult PMPost(List<ItemShowModel> collectedData, string SessionId, int cid)
        {
            try
            {
                Log("PM Post Method");
                #region Header
                string Message = "";
                Hashtable header = new Hashtable();
                header = new Hashtable
                        {
                           // { "DocNo", VoucherExistsVoucher},
                           // { "Date", CIssue.Date },
                            //{ "CustomerAC__Id", selectedincharge },
                            { "Branch__Id", 4}


                        };
                #endregion
                #region body
                List<Hashtable> body = new List<Hashtable>();
                Hashtable row1 = new Hashtable();
                foreach (var item in collectedData)
                {
                    // int Idate = ClsDataAcceslayer.GetDateToInt(Convert.ToDateTime(item.DueDate));
                    row1 = new Hashtable
                                      {
                                           {"Warehouse__Id",item.Whid},
                                           {"Item__Id", item.Itemid },
                                           {"Unit__Id",item.UnitId  },           //item.Units                                
                                           {"Quantity", item.QtytobeProduced },
                                           //{"AvailableStock", item.AvailableStockQty},
                                           //{"SafetyLevelQty", item.SafetyLevelQty },
                                           //{"Difference", item.Difference },
                                           //{"QuantitytobeProduced", item.QtytobeProduced },
                                           //{"TankTopUpCapacity", item.TopUpTankCapacity },
                                           //{"DueDate", Idate }
                                          // {"SafetyLevelQty", item.SafetyLevelQty },
                    };
                    body.Add(row1);
                }
                #endregion
                var postingData = new PostingData();
                postingData.data.Add(new Hashtable { { "Header", header }, { "Body", body } });
                string sContent = JsonConvert.SerializeObject(postingData);
                Log("Posting Updating sContent" + sContent);
                Log("Before Response");
                Log("Api URL:" + ProdPlan);
                var response = Focus8API.Post(PurInd, sContent, SessionId, ref error);
                Log("response:" + response);
                Log("After Response");
                if (response != null)
                {
                    var responseData = JsonConvert.DeserializeObject<APIResponse.PostResponse>(response);
                    if (responseData.result == -1)
                    {
                        Message = $"Posting Failed : {responseData.message } \n";
                        Log("Error posting failed:" + Message);
                    }
                    else
                    {
                        var docNo = Convert.ToString(responseData.data[0]["VoucherNo"]);
                        Message = "Posted into Purchase Indent successfully";
                    }
                }
                return Json(new { status = true, Message });
            }
            catch (Exception x)
            {
                return Json(new { status = false, Message = x.Message });
            }
        }
        //SFGExcel
        public ActionResult SFGExcel(List<ItemShowModel> collectedData)
        {
            try
            {
                Log("SFG Excel Method ");
                string Message = "";
                //var aCode = 65;
                //var workbook1 = new XLWorkbook();
                //workbook1.AddWorksheet("SFG Reorderlevel Report");
                //var worksheet = workbook1.Worksheet("SFG Reorderlevel Report");
                using (var workbook = new XLWorkbook())
                {
                    var worksheet = workbook.AddWorksheet("SFG Reorderlevel Report");
                    worksheet.Cell("A1").Value = "Reorder Level based on consumption of SFG";
                    var range = worksheet.Range("A1:O1");
                    range.Merge().Style.Font.SetBold().Font.FontSize = 16;
                    worksheet.Range("A1:O1").Row(1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    var currentRow = 2;
                    worksheet.Cell(currentRow, 1).Value = "Item";
                    worksheet.Cell(currentRow, 2).Value = "Units";
                    worksheet.Cell(currentRow, 3).Value = "Description";
                    worksheet.Cell(currentRow, 4).Value = "Available Stock Qty";
                    worksheet.Cell(currentRow, 5).Value = "Safety Level Qty";
                    worksheet.Cell(currentRow, 6).Value = "Difference";
                    worksheet.Cell(currentRow, 7).Value = "Qty to be produced";
                    worksheet.Cell(currentRow, 8).Value = "Tank Master";
                    worksheet.Cell(currentRow, 9).Value = "Tank Capacity";
                    worksheet.Cell(currentRow, 10).Value = "Closing Stock";
                    worksheet.Cell(currentRow, 11).Value = "Top up Tank Capacity";
                    worksheet.Cell(currentRow, 12).Value = "Production Planning Status";
                    worksheet.Cell(currentRow, 13).Value = "Production Planning DocNo";
                    worksheet.Cell(currentRow, 14).Value = "Due Date";
                    worksheet.Cell(currentRow, 15).Value = "Remarks";
                    //worksheet.Cell(currentRow, 16).Value = "Salesman Name";
                    for (int i = 0; i < collectedData.Count; i++)
                    {
                        var Items = collectedData[i];
                        currentRow++;
                        worksheet.Cell(currentRow, 1).SetValue(Items.Item == null ? "" : Items.Item);
                        worksheet.Cell(currentRow, 2).SetValue(Items.Units == null ? "" : Items.Units);
                        worksheet.Cell(currentRow, 3).SetValue(Items.Description == null ? "" : Items.Description);
                        worksheet.Cell(currentRow, 4).SetValue(Items.AvailableStockQty == 0 ? "" : Items.AvailableStockQty.ToString());
                        worksheet.Cell(currentRow, 5).SetValue(Items.SafetyLevelQty == 0 ? "" : Items.SafetyLevelQty.ToString());
                        worksheet.Cell(currentRow, 6).SetValue(Items.Difference == 0 ? "" : Items.Difference.ToString());
                        worksheet.Cell(currentRow, 7).SetValue(Items.QtytobeProduced);
                        worksheet.Cell(currentRow, 8).SetValue(Items.TankMaster);
                        worksheet.Cell(currentRow, 9).SetValue(Items.TankCapacity);
                        worksheet.Cell(currentRow, 10).SetValue(Items.ClosingStock);
                        worksheet.Cell(currentRow, 11).SetValue(Items.TopUpTankCapacity);
                        worksheet.Cell(currentRow, 12).SetValue(Items.ProductionPlanningStatus == null ? "" : Items.ProductionPlanningStatus);
                        worksheet.Cell(currentRow, 13).SetValue(Items.ProductionPlanningDocNo == null ? "" : Items.ProductionPlanningDocNo);
                        worksheet.Cell(currentRow, 14).SetValue(Items.DueDate);
                        worksheet.Cell(currentRow, 15).SetValue(Items.Remarks == null ? "" : Items.Remarks);
                    }
                    worksheet.Columns().AdjustToContents();
                    worksheet.Range($@"A2:O{currentRow}").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    worksheet.Range($@"A2:O{currentRow}").Style.Border.TopBorder = XLBorderStyleValues.Thin;
                    worksheet.Range($@"A2:O{currentRow}").Style.Border.InsideBorder = XLBorderStyleValues.Thin;
                    worksheet.Range($@"A2:O{currentRow}").Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                    worksheet.Range($@"A2:O{currentRow}").Style.Border.LeftBorder = XLBorderStyleValues.Thin;
                    worksheet.Range($@"A2:O{currentRow}").Style.Border.RightBorder = XLBorderStyleValues.Thin;
                    var rango = worksheet.Range("A2:O2");
                    rango.Style.Font.FontSize = 12; //Indicamos el tamaño de la fuente
                    rango.Style.Font.FontColor = XLColor.White;
                    rango.Style.Fill.BackgroundColor = XLColor.FromHtml("#0073AA");
                    using (MemoryStream stream = new MemoryStream())
                    {
                        workbook.SaveAs(stream);


                        worksheet.Columns("A:AZ").AdjustToContents();
                        var fileName = $"{"Re Order Level SFG Report" + DateTime.Now.Date.ToString("ddMMyyyy")}.xlsx";
                        string tempPath = Server.MapPath("~/Temp");
                        var dirInfo = new DirectoryInfo(tempPath);
                        if (!dirInfo.Exists)
                        {
                            dirInfo.Create();
                        }
                        var savePath = Path.Combine(tempPath, fileName);
                        workbook.SaveAs(savePath);

                        return Json(new { status = true, fileName });

                    }
                }
                //return Json(new { status = true, Message });
            }
            catch (Exception x)
            {
                return Json(new { status = false, Message = x.Message });
            }
        }
        public ActionResult FGExcel(List<ItemShowModel> collectedData)
        {
            try
            {
                // Log("FG Excel Method");
                //string Message = "";
                Log("Exporting Excel FG Entered..");
                //var aCode = 65;
                //var workbook1 = new XLWorkbook();
                //workbook1.AddWorksheet("SFG Reorderlevel Report");
                //var worksheet = workbook1.Worksheet("SFG Reorderlevel Report");
                using (var workbook = new XLWorkbook())
                {
                    var worksheet = workbook.AddWorksheet("FG Reorderlevel Report");
                    worksheet.Cell("A1").Value = "Reorder Level based on consumption of FG";
                    var range = worksheet.Range("A1:J1");
                    range.Merge().Style.Font.SetBold().Font.FontSize = 16;
                    worksheet.Range("A1:J1").Row(1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    var currentRow = 2;
                    worksheet.Cell(currentRow, 1).Value = "Item";
                    worksheet.Cell(currentRow, 2).Value = "Units";
                    worksheet.Cell(currentRow, 3).Value = "Description";
                    worksheet.Cell(currentRow, 4).Value = "Available Stock Qty";
                    worksheet.Cell(currentRow, 5).Value = "Safety Level Qty";
                    worksheet.Cell(currentRow, 6).Value = "Difference";
                    worksheet.Cell(currentRow, 7).Value = "Qty to be produced";
                    //worksheet.Cell(currentRow, 8).Value = "Tank Master";
                    // worksheet.Cell(currentRow, 9).Value = "Tank Capacity";
                    //worksheet.Cell(currentRow, 10).Value = "Closing Stock";
                    //worksheet.Cell(currentRow, 11).Value = "Top up Tank Capacity";
                    worksheet.Cell(currentRow, 8).Value = "Production Planning Status";
                    worksheet.Cell(currentRow, 9).Value = "Production Planning DocNo";
                    worksheet.Cell(currentRow, 10).Value = "Due Date";
                    //worksheet.Cell(currentRow, 15).Value = "Remarks";
                    //worksheet.Cell(currentRow, 16).Value = "Salesman Name";
                    for (int i = 0; i < collectedData.Count; i++)
                    {
                        var Items = collectedData[i];
                        currentRow++;
                        worksheet.Cell(currentRow, 1).SetValue(Items.Item);
                        worksheet.Cell(currentRow, 2).SetValue(Items.Units == null ? "" : Items.Units);
                        worksheet.Cell(currentRow, 3).SetValue(Items.Description == null ? "" : Items.Description);
                        worksheet.Cell(currentRow, 4).SetValue(Items.AvailableStockQty);
                        worksheet.Cell(currentRow, 5).SetValue(Items.SafetyLevelQty);
                        worksheet.Cell(currentRow, 6).SetValue(Items.Difference);
                        worksheet.Cell(currentRow, 7).SetValue(Items.QtytobeProduced);
                        // worksheet.Cell(currentRow, 8).SetValue(Items.TankMaster);
                        // worksheet.Cell(currentRow, 9).SetValue(Items.TankCapacity);
                        // worksheet.Cell(currentRow, 10).SetValue(Items.ClosingStock);
                        // worksheet.Cell(currentRow, 11).SetValue(Items.TopUpTankCapacity);
                        worksheet.Cell(currentRow, 8).SetValue(Items.ProductionPlanningStatus == null ? "" : Items.ProductionPlanningStatus);
                        worksheet.Cell(currentRow, 9).SetValue(Items.ProductionPlanningDocNo == null ? "" : Items.ProductionPlanningDocNo);
                        worksheet.Cell(currentRow, 10).SetValue(Items.DueDate);
                        //worksheet.Cell(currentRow, 15).SetValue(Items.Remarks == null ? "" : Items.Remarks);
                    }
                    worksheet.Columns().AdjustToContents();
                    worksheet.Range($@"A2:J{currentRow}").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    worksheet.Range($@"A2:J{currentRow}").Style.Border.TopBorder = XLBorderStyleValues.Thin;
                    worksheet.Range($@"A2:J{currentRow}").Style.Border.InsideBorder = XLBorderStyleValues.Thin;
                    worksheet.Range($@"A2:J{currentRow}").Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                    worksheet.Range($@"A2:J{currentRow}").Style.Border.LeftBorder = XLBorderStyleValues.Thin;
                    worksheet.Range($@"A2:J{currentRow}").Style.Border.RightBorder = XLBorderStyleValues.Thin;
                    var rango = worksheet.Range("A2:J2");
                    rango.Style.Font.FontSize = 12; //Indicamos el tamaño de la fuente
                    rango.Style.Font.FontColor = XLColor.White;
                    rango.Style.Fill.BackgroundColor = XLColor.FromHtml("#0073AA");
                    using (MemoryStream stream = new MemoryStream())
                    {
                        workbook.SaveAs(stream);


                        worksheet.Columns("A:AZ").AdjustToContents();
                        var fileName = $"{"Re Order Level FG Report" + DateTime.Now.Date.ToString("ddMMyyyy")}.xlsx";
                        string tempPath = Server.MapPath("~/Temp");
                        var dirInfo = new DirectoryInfo(tempPath);
                        if (!dirInfo.Exists)
                        {
                            dirInfo.Create();
                        }
                        var savePath = Path.Combine(tempPath, fileName);
                        workbook.SaveAs(savePath);
                        Log("Excel for Fg path: " + savePath);
                        return Json(new { status = true, fileName });

                    }
                }
                //return Json(new { status = true, Message });
            }
            catch (Exception x)
            {
                return Json(new { status = false, Message = x.Message });
            }
        }
        //RMBaseOilsExcel
        public ActionResult RMBaseOilsExcel(List<ItemShowModel> collectedData)
        {
            try
            {

                Log("Exporting Excel BaseOils Entered..");
                //var aCode = 65;
                //var workbook1 = new XLWorkbook();
                //workbook1.AddWorksheet("SFG Reorderlevel Report");
                //var worksheet = workbook1.Worksheet("SFG Reorderlevel Report");
                using (var workbook = new XLWorkbook())
                {
                    var worksheet = workbook.AddWorksheet("BaseOils Reorderlevel Report");
                    worksheet.Cell("A1").Value = "Reorder Level based on consumption of BaseOils";
                    var range = worksheet.Range("A1:O1");
                    range.Merge().Style.Font.SetBold().Font.FontSize = 16;
                    worksheet.Range("A1:O1").Row(1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    var currentRow = 2;
                    worksheet.Cell(currentRow, 1).Value = "Item";
                    worksheet.Cell(currentRow, 2).Value = "Units";
                    worksheet.Cell(currentRow, 3).Value = "Description";
                    worksheet.Cell(currentRow, 4).Value = "Available Stock Qty";
                    worksheet.Cell(currentRow, 5).Value = "Safety Level Qty";
                    worksheet.Cell(currentRow, 6).Value = "Difference";
                    worksheet.Cell(currentRow, 7).Value = "Qty to be produced";
                    worksheet.Cell(currentRow, 8).Value = "Tank Master";
                    worksheet.Cell(currentRow, 9).Value = "Tank Capacity";
                    worksheet.Cell(currentRow, 10).Value = "Closing Stock";
                    worksheet.Cell(currentRow, 11).Value = "Available RM Tank Capacity";
                    worksheet.Cell(currentRow, 12).Value = "Purchase Indent Status";
                    worksheet.Cell(currentRow, 13).Value = "Purchase Indent DocNo";
                    worksheet.Cell(currentRow, 14).Value = "Due Date";
                    worksheet.Cell(currentRow, 15).Value = "Warehouse";
                    //worksheet.Cell(currentRow, 15).Value = "Remarks";
                    //worksheet.Cell(currentRow, 16).Value = "Salesman Name";
                    for (int i = 0; i < collectedData.Count; i++)
                    {
                        var Items = collectedData[i];
                        currentRow++;
                        worksheet.Cell(currentRow, 1).SetValue(Items.Item == null ? "" : Items.Item);
                        worksheet.Cell(currentRow, 2).SetValue(Items.Units == null ? "" : Items.Units);
                        worksheet.Cell(currentRow, 3).SetValue(Items.Description == null ? "" : Items.Description);
                        worksheet.Cell(currentRow, 4).SetValue(Items.AvailableStockQty == 0 ? "" : Items.AvailableStockQty.ToString());
                        worksheet.Cell(currentRow, 5).SetValue(Items.SafetyLevelQty == 0 ? "" : Items.SafetyLevelQty.ToString());
                        worksheet.Cell(currentRow, 6).SetValue(Items.Difference == 0 ? "" : Items.Difference.ToString());
                        worksheet.Cell(currentRow, 7).SetValue(Items.QtytobeProduced == 0 ? "" : Items.QtytobeProduced.ToString());
                        worksheet.Cell(currentRow, 8).SetValue(Items.TankMaster);
                        worksheet.Cell(currentRow, 9).SetValue(Items.TankCapacity);
                        worksheet.Cell(currentRow, 10).SetValue(Items.ClosingStock);
                        worksheet.Cell(currentRow, 11).SetValue(Items.TopUpTankCapacity);
                        worksheet.Cell(currentRow, 12).SetValue(Items.ProductionPlanningStatus == null ? "" : Items.ProductionPlanningStatus);
                        worksheet.Cell(currentRow, 13).SetValue(Items.ProductionPlanningDocNo == null ? "" : Items.ProductionPlanningDocNo);
                        worksheet.Cell(currentRow, 14).SetValue(Items.DueDate);
                        worksheet.Cell(currentRow, 15).SetValue(Items.Warehouse == null ? "" : Items.Warehouse);
                        //worksheet.Cell(currentRow, 15).SetValue(Items.Remarks == null ? "" : Items.Remarks);
                    }
                    worksheet.Columns().AdjustToContents();
                    worksheet.Range($@"A2:O{currentRow}").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    worksheet.Range($@"A2:O{currentRow}").Style.Border.TopBorder = XLBorderStyleValues.Thin;
                    worksheet.Range($@"A2:O{currentRow}").Style.Border.InsideBorder = XLBorderStyleValues.Thin;
                    worksheet.Range($@"A2:O{currentRow}").Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                    worksheet.Range($@"A2:O{currentRow}").Style.Border.LeftBorder = XLBorderStyleValues.Thin;
                    worksheet.Range($@"A2:O{currentRow}").Style.Border.RightBorder = XLBorderStyleValues.Thin;
                    var rango = worksheet.Range("A2:O2");
                    rango.Style.Font.FontSize = 12; //Indicamos el tamaño de la fuente
                    rango.Style.Font.FontColor = XLColor.White;
                    rango.Style.Fill.BackgroundColor = XLColor.FromHtml("#0073AA");
                    using (MemoryStream stream = new MemoryStream())
                    {
                        workbook.SaveAs(stream);


                        worksheet.Columns("A:AZ").AdjustToContents();
                        var fileName = $"{"ReOrder Level BaseOils Report" + DateTime.Now.Date.ToString("ddMMyyyy")}.xlsx";
                        string tempPath = Server.MapPath("~/Temp");
                        var dirInfo = new DirectoryInfo(tempPath);
                        if (!dirInfo.Exists)
                        {
                            dirInfo.Create();
                        }
                        var savePath = Path.Combine(tempPath, fileName);
                        workbook.SaveAs(savePath);
                        Log("Excel for Fg path: " + savePath);
                        return Json(new { status = true, fileName });
                    }
                }
            }
            catch (Exception x)
            {
                return Json(new { status = false, Message = x.Message });
            }
        }
        //RMAdditiveExcel
        public ActionResult RMAdditiveExcel(List<ItemShowModel> collectedData)
        {
            try
            {
                Log("Exporting Excel Additives Entered..");
                //var aCode = 65;
                //var workbook1 = new XLWorkbook();
                //workbook1.AddWorksheet("SFG Reorderlevel Report");
                //var worksheet = workbook1.Worksheet("SFG Reorderlevel Report");
                using (var workbook = new XLWorkbook())
                {
                    var worksheet = workbook.AddWorksheet("Additives Reorderlevel Report");
                    worksheet.Cell("A1").Value = "Reorder Level based on consumption of Additives";
                    var range = worksheet.Range("A1:K1");
                    range.Merge().Style.Font.SetBold().Font.FontSize = 16;
                    worksheet.Range("A1:K1").Row(1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    var currentRow = 2;
                    worksheet.Cell(currentRow, 1).Value = "Item";
                    worksheet.Cell(currentRow, 2).Value = "Warehouse";
                    worksheet.Cell(currentRow, 3).Value = "Units";
                    worksheet.Cell(currentRow, 4).Value = "Description";
                    worksheet.Cell(currentRow, 5).Value = "Available Stock Qty";
                    worksheet.Cell(currentRow, 6).Value = "Safety Level Qty";
                    worksheet.Cell(currentRow, 7).Value = "Difference";
                    worksheet.Cell(currentRow, 8).Value = "Qty to be produced";
                    //worksheet.Cell(currentRow, 8).Value = "Tank Master";
                    // worksheet.Cell(currentRow, 9).Value = "Tank Capacity";
                    //worksheet.Cell(currentRow, 10).Value = "Closing Stock";
                    //worksheet.Cell(currentRow, 11).Value = "Top up Tank Capacity";
                    worksheet.Cell(currentRow, 9).Value = "Purchase Indent Status";
                    worksheet.Cell(currentRow, 10).Value = "Purchase Indent DocNo";
                    worksheet.Cell(currentRow, 11).Value = "Due Date";
                    //worksheet.Cell(currentRow, 15).Value = "Remarks";
                    //worksheet.Cell(currentRow, 16).Value = "Salesman Name";
                    for (int i = 0; i < collectedData.Count; i++)
                    {
                        var Items = collectedData[i];
                        currentRow++;
                        worksheet.Cell(currentRow, 1).SetValue(Items.Item == null ? "" : Items.Item);
                        worksheet.Cell(currentRow, 2).SetValue(Items.Warehouse == null ? "" : Items.Warehouse);
                        worksheet.Cell(currentRow, 3).SetValue(Items.Units == null ? "" : Items.Units);
                        worksheet.Cell(currentRow, 4).SetValue(Items.Description == null ? "" : Items.Description);
                        worksheet.Cell(currentRow, 5).SetValue(Items.AvailableStockQty == 0 ? "" : Items.AvailableStockQty.ToString());
                        worksheet.Cell(currentRow, 6).SetValue(Items.SafetyLevelQty == 0 ? "" : Items.SafetyLevelQty.ToString());
                        worksheet.Cell(currentRow, 7).SetValue(Items.Difference == 0 ? "" : Items.Difference.ToString());
                        worksheet.Cell(currentRow, 8).SetValue(Items.QtytobeProduced);
                        // worksheet.Cell(currentRow, 8).SetValue(Items.TankMaster);
                        // worksheet.Cell(currentRow, 9).SetValue(Items.TankCapacity);
                        // worksheet.Cell(currentRow, 10).SetValue(Items.ClosingStock);
                        // worksheet.Cell(currentRow, 11).SetValue(Items.TopUpTankCapacity);
                        worksheet.Cell(currentRow, 9).SetValue(Items.ProductionPlanningStatus == null ? "" : Items.ProductionPlanningStatus);
                        worksheet.Cell(currentRow, 10).SetValue(Items.ProductionPlanningDocNo == null ? "" : Items.ProductionPlanningDocNo);
                        worksheet.Cell(currentRow, 11).SetValue(Items.DueDate);
                        //worksheet.Cell(currentRow, 15).SetValue(Items.Remarks == null ? "" : Items.Remarks);
                    }
                    worksheet.Columns().AdjustToContents();
                    worksheet.Range($@"A2:K{currentRow}").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    worksheet.Range($@"A2:K{currentRow}").Style.Border.TopBorder = XLBorderStyleValues.Thin;
                    worksheet.Range($@"A2:K{currentRow}").Style.Border.InsideBorder = XLBorderStyleValues.Thin;
                    worksheet.Range($@"A2:K{currentRow}").Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                    worksheet.Range($@"A2:K{currentRow}").Style.Border.LeftBorder = XLBorderStyleValues.Thin;
                    worksheet.Range($@"A2:K{currentRow}").Style.Border.RightBorder = XLBorderStyleValues.Thin;
                    var rango = worksheet.Range("A2:K2");
                    rango.Style.Font.FontSize = 12; //Indicamos el tamaño de la fuente
                    rango.Style.Font.FontColor = XLColor.White;
                    rango.Style.Fill.BackgroundColor = XLColor.FromHtml("#0073AA");
                    using (MemoryStream stream = new MemoryStream())
                    {
                        workbook.SaveAs(stream);


                        worksheet.Columns("A:AZ").AdjustToContents();
                        var fileName = $"{"ReOrder Level Addtives Report" + DateTime.Now.Date.ToString("ddMMyyyy")}.xlsx";
                        string tempPath = Server.MapPath("~/Temp");
                        var dirInfo = new DirectoryInfo(tempPath);
                        if (!dirInfo.Exists)
                        {
                            dirInfo.Create();
                        }
                        var savePath = Path.Combine(tempPath, fileName);
                        workbook.SaveAs(savePath);
                        Log("Excel for Fg path: " + savePath);
                        return Json(new { status = true, fileName });

                    }
                }
            }
            catch (Exception x)
            {
                return Json(new { status = false, Message = x.Message });
            }
        }
        //PMExcel
        public ActionResult PMExcel(List<ItemShowModel> collectedData)
        {
            try
            {
                Log("Exporting Excel PM Entered..");
                //var aCode = 65;
                //var workbook1 = new XLWorkbook();
                //workbook1.AddWorksheet("SFG Reorderlevel Report");
                //var worksheet = workbook1.Worksheet("SFG Reorderlevel Report");
                using (var workbook = new XLWorkbook())
                {
                    var worksheet = workbook.AddWorksheet("PM Reorderlevel Report");
                    worksheet.Cell("A1").Value = "Reorder Level based on consumption of PM";
                    var range = worksheet.Range("A1:K1");
                    range.Merge().Style.Font.SetBold().Font.FontSize = 16;
                    worksheet.Range("A1:K1").Row(1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    var currentRow = 2;
                    worksheet.Cell(currentRow, 1).Value = "Item";
                    worksheet.Cell(currentRow, 2).Value = "Warehouse";
                    worksheet.Cell(currentRow, 3).Value = "Units";
                    worksheet.Cell(currentRow, 4).Value = "Description";
                    worksheet.Cell(currentRow, 5).Value = "Available Stock Qty";
                    worksheet.Cell(currentRow, 6).Value = "Safety Level Qty";
                    worksheet.Cell(currentRow, 7).Value = "Difference";
                    worksheet.Cell(currentRow, 8).Value = "Qty to be produced";
                    //worksheet.Cell(currentRow, 8).Value = "Tank Master";
                    // worksheet.Cell(currentRow, 9).Value = "Tank Capacity";
                    //worksheet.Cell(currentRow, 10).Value = "Closing Stock";
                    //worksheet.Cell(currentRow, 11).Value = "Top up Tank Capacity";
                    worksheet.Cell(currentRow, 9).Value = "Purchase Indent Status";
                    worksheet.Cell(currentRow, 10).Value = "Purchase Indent DocNo";
                    worksheet.Cell(currentRow, 11).Value = "Due Date";
                    //worksheet.Cell(currentRow, 15).Value = "Remarks";
                    //worksheet.Cell(currentRow, 16).Value = "Salesman Name";
                    for (int i = 0; i < collectedData.Count; i++)
                    {
                        var Items = collectedData[i];
                        currentRow++;
                        worksheet.Cell(currentRow, 1).SetValue(Items.Item == null ? "" : Items.Item);
                        worksheet.Cell(currentRow, 2).SetValue(Items.Warehouse == null ? "" : Items.Warehouse);
                        worksheet.Cell(currentRow, 3).SetValue(Items.Units == null ? "" : Items.Units);
                        worksheet.Cell(currentRow, 4).SetValue(Items.Description == null ? "" : Items.Description);
                        worksheet.Cell(currentRow, 5).SetValue(Items.AvailableStockQty == 0 ? "" : Items.AvailableStockQty.ToString());
                        worksheet.Cell(currentRow, 6).SetValue(Items.SafetyLevelQty);
                        worksheet.Cell(currentRow, 7).SetValue(Items.Difference);
                        worksheet.Cell(currentRow, 8).SetValue(Items.QtytobeProduced);
                        // worksheet.Cell(currentRow, 8).SetValue(Items.TankMaster);
                        // worksheet.Cell(currentRow, 9).SetValue(Items.TankCapacity);
                        // worksheet.Cell(currentRow, 10).SetValue(Items.ClosingStock);
                        // worksheet.Cell(currentRow, 11).SetValue(Items.TopUpTankCapacity);
                        worksheet.Cell(currentRow, 9).SetValue(Items.ProductionPlanningStatus == null ? "" : Items.ProductionPlanningStatus);
                        worksheet.Cell(currentRow, 10).SetValue(Items.ProductionPlanningDocNo == null ? "" : Items.ProductionPlanningDocNo);
                        worksheet.Cell(currentRow, 11).SetValue(Items.DueDate);
                        //worksheet.Cell(currentRow, 15).SetValue(Items.Remarks == null ? "" : Items.Remarks);
                    }
                    worksheet.Columns().AdjustToContents();
                    worksheet.Range($@"A2:K{currentRow}").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    worksheet.Range($@"A2:K{currentRow}").Style.Border.TopBorder = XLBorderStyleValues.Thin;
                    worksheet.Range($@"A2:K{currentRow}").Style.Border.InsideBorder = XLBorderStyleValues.Thin;
                    worksheet.Range($@"A2:K{currentRow}").Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                    worksheet.Range($@"A2:K{currentRow}").Style.Border.LeftBorder = XLBorderStyleValues.Thin;
                    worksheet.Range($@"A2:K{currentRow}").Style.Border.RightBorder = XLBorderStyleValues.Thin;
                    var rango = worksheet.Range("A2:K2");
                    rango.Style.Font.FontSize = 12; //Indicamos el tamaño de la fuente
                    rango.Style.Font.FontColor = XLColor.White;
                    rango.Style.Fill.BackgroundColor = XLColor.FromHtml("#0073AA");
                    using (MemoryStream stream = new MemoryStream())
                    {
                        workbook.SaveAs(stream);


                        worksheet.Columns("A:AZ").AdjustToContents();
                        var fileName = $"{"ReOrder Level PM Report" + DateTime.Now.Date.ToString("ddMMyyyy")}.xlsx";
                        string tempPath = Server.MapPath("~/Temp");
                        var dirInfo = new DirectoryInfo(tempPath);
                        if (!dirInfo.Exists)
                        {
                            dirInfo.Create();
                        }
                        var savePath = Path.Combine(tempPath, fileName);
                        workbook.SaveAs(savePath);
                        Log("Excel for Fg path: " + savePath);
                        return Json(new { status = true, fileName });

                    }
                }
            }
            catch (Exception x)
            {
                return Json(new { status = false, Message = x.Message });
            }
        }
        [HttpGet]
        [DeleteFileAttribute]
        public ActionResult Download1(string file)
        {
            string fullPath = System.IO.Path.Combine(Server.MapPath("~/Temp"), file);
            return File(fullPath, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", file);
        }


        internal class DeleteFileAttribute : ActionFilterAttribute
        {
            public override void OnResultExecuted(ResultExecutedContext filterContext)
            {
                filterContext.HttpContext.Response.Flush();
                string filePath = (filterContext.Result as FilePathResult).FileName;
                System.IO.File.Delete(filePath);
            }
        }
        //GetPpl
        public ActionResult GetPpl(int CompanyId, int vtype, string DocNo, int UserId)
        {
            try
            {

                Log("Get Production Planning Entered");
                string Message = "";
                int bodyid = 0;
                int item; int n = 0; int unit; int bom; int duedate; decimal qty; int version; string itemname = ""; string ii = ""; Boolean status = true;
                string qgetppl = $@"SELECT sVoucherNo,ind.iProduct,mp.sName Item,mu.iMasterId UnitId,QuantitytobeProduced,
	                                 DueDate,iBOM,isnull(iVersion,0) iVersion,d.iBodyId FROM --convert(varchar,dbo.IntToDate(DueDate),103)
	                                tCore_Header_0 h JOIN tCore_Data_0 d on h.iHeaderId=d.iHeaderId
	                                join tCore_Indta_0 ind on ind.iBodyId=d.iBodyId
	                                join tCore_Data{vtype}_0 dv on dv.iBodyId=d.iBodyId
	                                join mCore_Units mu on mu.iMasterId=iUnit
	                                join muCore_Product_Replenishment mur on mur.iMasterId=iProduct
                                    join mCore_Product mp on mp.iMasterId=mur.iMasterId
	                                left join mMRP_BomVariantHeader bh on bh.iVariantId = mur.iBOM
	                                left join mMRP_BomHeader bom on bom.iBomId=bh.iBomId
	                                where iVoucherType={vtype} and sVoucherNo='{DocNo}'";
                Log("" + qgetppl);
                DataSet dsgetppl = ClsDataAcceslayer.GetData1(qgetppl, CompanyId, ref error);
                Log("query result count:" + dsgetppl.Tables[0].Rows.Count);
                if (dsgetppl != null && dsgetppl.Tables[0].Rows.Count > 0)
                {
                    for (int i = 0; i < dsgetppl.Tables[0].Rows.Count; i++)
                    {
                        item = Convert.ToInt32(dsgetppl.Tables[0].Rows[i]["iProduct"]);
                        itemname = dsgetppl.Tables[0].Rows[i]["Item"].ToString();
                        unit = Convert.ToInt32(dsgetppl.Tables[0].Rows[i]["UnitId"]);
                        qty = Convert.ToDecimal(dsgetppl.Tables[0].Rows[i]["QuantitytobeProduced"]);
                        duedate = Convert.ToInt32(dsgetppl.Tables[0].Rows[i]["DueDate"]);
                        bom = Convert.ToInt32(dsgetppl.Tables[0].Rows[i]["iBOM"]);
                        version = Convert.ToInt32(dsgetppl.Tables[0].Rows[i]["iVersion"]);//==null?0: dsgetppl.Tables[0].Rows[i]["iVersion"]
                        bodyid = Convert.ToInt32(dsgetppl.Tables[0].Rows[i]["iBodyId"]);
                        Log($"Item :{item},Unit :{unit},Quantity :{qty},DueDate :{duedate},Bom:{bom},Version :{version}");
                        if (bom == 0)
                        {
                            n++;
                            // ii = ii+ itemname + ",";
                            Message += $"{itemname} is not mapped with BOM,Production Order is not raised for it.\n";//Item(s) 

                            Log($"{itemname} is not mapped with BOM,Production Order is not raised for it");
                            // return Json(new { status = false, Message = "Item is not mapped with BOM" });
                        }
                        else
                        { //}
                            string qpordorder0 = $@"INSERT INTO tMrp_ProdOrder_0 
                       (iProdOrderId,sProdOrderNo, iDate, iDueDate, sRemarks, iOrderStatus, iWareHouseId, iVaraintId,
                        iIssueType, sSONO, iCustomer, sSpecialInstruction,sRefOrderNo,iType,iPDRID,iMainProdId,
                        iCreatedDate,iCreatedTime,iModifiedDate,iModifiedTime,iCreatedBy,iModifiedBy,iTagFilterId,
                        iTagFilterValue,sBatchNo,iSalesOrderNo) 
                      VALUES (
                      (select isnull(MAX(iProdOrderId),0)+1 from tMrp_ProdOrder_0)
                      ,(SELECT TOP 1 
                            CASE 
                                WHEN COUNT(*) = 0 THEN 'PRO/' + '1'  -- If the table is empty
                                ELSE 
                                    MAX(SUBSTRING(sProdOrderNo, 1, CHARINDEX('/', sProdOrderNo, 1))) +
                                    CAST((ISNULL(MAX(iProdOrderId), 0) + 1) AS VARCHAR)
                            END
                        FROM 
                            tMrp_ProdOrder_0), 
                      dbo.datetoint(getdate()), {duedate}, '', 2,0, {bom}, 0, '', 0, '', '', 0, 0,
                       0,dbo.datetoint(getdate()),dbo.fCore_TimeToInt(CONVERT(VARCHAR, GETDATE(), 108)),0,0,{UserId},-1,0,0,'',0 )";
                            Log("Query tMrp_ProdOrder_0:" + qpordorder0);
                            int n1 = obj.GetQueryExe(qpordorder0, CompanyId, ref error);  //iProdOrderId   5, 
                            int n2 = 0;

                            Log("No of rows effected for tMrp_ProdOrder_0" + n1);
                            if (n1 > 0)
                            {
                                string qprodorderbody0 = $@"INSERT INTO tMrp_ProdOrderBody_0 (
                              iProdOrderId, iItem, iUnit, fQuantity)
                            VALUES( (select iProdOrderId from tMrp_ProdOrder_0 where sProdOrderNo=(Select TOP 1 sProdOrderNo
                           FROM tMrp_ProdOrder_0 Order by LEN(sProdOrderNo) desc,sProdOrderNo desc)), {item}, {unit}, {qty})";
                                n2 = obj.GetQueryExe(qprodorderbody0, CompanyId, ref error);
                                Log("Query for tMrp_ProdOrderBody_0 :" + qprodorderbody0);
                                Log("No. of rows effected for tMrp_ProdOrderBody_0:" + n2);
                            }
                            if (n1 > 0 && n2 > 0)
                            {
                                string qupdate = $@"update tCore_Data{vtype}_0 set ProductionOrderNo_= (Select TOP 1 sProdOrderNo FROM tMrp_ProdOrder_0 Order by LEN(sProdOrderNo) desc,sProdOrderNo desc)
from tCore_Header_0 h join tCore_Data_0 d on h.iHeaderId=d.iHeaderId
join tCore_Indta_0 ind on ind.iBodyId=d.iBodyId
join tCore_Data{vtype}_0 dv on dv.iBodyId=d.iBodyId
where sVoucherNo='{DocNo}' and iVoucherType={vtype} and iProduct={item} and QuantitytobeProduced={qty} and d.iBodyId={bodyid}";
                                Log("Update ProductionOrderNo query :" + qupdate);
                                int n3 = obj.GetExecute(qupdate, CompanyId, ref error);
                                Log("No.of rows effected:" + n3);
                                Message += $"Production order raised successfully for item {itemname}.\n";
                            }
                            else
                            {
                                status = false;
                                Log("Error :" + error);
                                Message += "Issue :" + error + ".\n";
                            }
                        }
                    }
                }
                if (n == dsgetppl.Tables[0].Rows.Count) { status = false; }
                return Json(new { status = status, Message = Message });
            }
            catch (Exception x)
            {
                Log("Exception in GetPpl method :" + x.StackTrace);
                return Json(new { status = false, Message = x.Message });
            }
        }

        public static void Log(string content)
        {
            StreamWriter streamWriter = new StreamWriter((Stream)new FileStream(Path.GetTempPath() + DateTime.Now.ToString("ddMM") + "SagarPetrolPvtLtd.log", FileMode.OpenOrCreate, FileAccess.Write));
            streamWriter.BaseStream.Seek(0L, SeekOrigin.End);
            streamWriter.WriteLine(System.DateTime.Now.ToShortDateString() + "-" + System.DateTime.Now.ToShortTimeString() + " - " + content);
            streamWriter.WriteLine();
            streamWriter.Flush();
            streamWriter.Close();
        }
    }
}