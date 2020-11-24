using AESEncryp;
using DBOper;
using Newtonsoft.Json;
using Newtonsoft.Json.Converters;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Dynamic;
using System.IO;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;
using System.Web.Services;

namespace SHWXWebSvr
{
    /// <summary>
    /// SHWXSvr 的摘要说明
    /// </summary>
    [WebService(Namespace = "https://**.**.com/")]
    [WebServiceBinding(ConformsTo = WsiProfiles.BasicProfile1_1)]
    [System.Web.Script.Services.ScriptService]  //允许前端调用
    [System.ComponentModel.ToolboxItem(false)]
    // 若要允许使用 ASP.NET AJAX 从脚本中调用此 Web 服务，请取消注释以下行。 
    // [System.Web.Script.Services.ScriptService]
    public class SHWXSvr : System.Web.Services.WebService
    {
        private const string cAppID = "****";
        private const string cAppSecret = "****";

        public class FormData
        {
            public string Name { get; set; }
            public string Value { get; set; }
        }

        public class TableData
        {
            public string Title { get; set; }
            //public string Field { get; set; }
            public object Table { get; set; }
        }

        public class AnnexData
        {
            public string FileName { get; set; }
            public string CreateMan { get; set; }
            public DateTime CreateDate { get; set; }
            public string Url { get; set; }
        }

        public class RelationData
        {
            public string Subject { get; set; }
            public string BillID { get; set; }
            public string FlowGUID { get; set; }
        }

        public class OpinionData
        {
            public string NodeName { get; set; }
            public string DoMan { get; set; }
            public string DoType { get; set; }
            public DateTime DoTime { get; set; }
            public string GetMan { get; set; }
            public string Opinion { get; set; }
        }

        public class FlowInfo
        {
            public List<FormData> formData;
            public List<TableData> tableData;
            public List<AnnexData> annexData;
            public List<RelationData> relationData;
            public List<OpinionData> opinionData;
            public bool isAllowRebut;
            public bool isAllowConsult;
            public string indexNodeName;
        }
        public class ConFlowData  //合同的相关流程
        {
            public string FlowGUID { get; set; } //流程GUID
            public string WorkflowName { get; set; }  //流程类型
            public string ProjectName { get; set; } //归属项目
            public string Subject { get; set; } //流程标题
            public string BillID { get; set; } //流程编号
            public DateTime CreateDate { get; set; } //创建日期
            public string CreateMan { get; set; } //创建人
            public string State { get; set; }  //状态
            public string NodeName { get; set; }  //当前节点
            public decimal TotalEffMoney { get; set; }  //总影响金额
            public decimal ShareEffMoney { get; set; }  //分摊影响金额
            public int RowNumber { get; set; }  //排序号（用这个排序，不用时间排序）
        }

        public class ConSuplyData  //合同的补充协议
        {
            public string ConGUID { get; set; }  //合同GUID
            public string ProjectName { get; set; }  //归属项目
            public string ConName { get; set; }  //合同名称
            public string ConNo { get; set; }    //合同编号
            public decimal ConMoney { get; set; } //合同金额
            public string ProvName { get; set; }  //供应商/乙方
            public DateTime SignDate { get; set; }  //签订日期
        }

        //合同详情类
        public class ConInfo
        {
            public List<FormData> formData;
            public List<AnnexData> annexData;
            public List<ConSuplyData> conSuplyData;
            public List<ConFlowData> conFlowData;
        }

        public class OrgManData  //通讯录结构
        {
            public string Name { get; set; }  //人or组织的名称
            public string Type { get; set; }  //人or组织
            public string ID { get; set; }    //人or组织的ID,人以MAN开头,组织以ORG开头
            public string PID { get; set; }   //人or组织的PID,人以MAN开头,组织以ORG开头
            public string Order { get; set; }    //字符排序
            public string Mobile { get; set; }
            public string Depart { get; set; }//部门/职位
            public string AvatarUrl { get; set; }
            public string UserID { get; set; }
            public string Gender { get; set; }
            public string isHide { get; set; }
        }

        public class PortalNews
        {
            public string Img { get; set; }
            public string Url { get; set; }
        }

        public class PortalData
        {
            public string ToDoCount { get; set; }
            public string MakeCopyCount { get; set; }
            public string CustCount1 { get; set; }
            public string CustCount2 { get; set; }
            public string RGAmount { get; set; }
            public string QYAmount { get; set; }
            public string HKAmount { get; set; }
        }

        public class UserInfo
        {
            public string UserGUID { get; set; }
            public string UserName { get; set; }
            public string UserID { get; set; }
            public string MobileNum { get; set; }
            public string AvatarUrl { get; set; }
            public string Power { get; set; }
            public string GXKLPower { get; set; }
            public string LSKLPower { get; set; }
        }

        public class RetObj
        {
            public string errNo;
            public string errDesc;
            public object retVal;
        }

        public static string PostHttpData(string url)
        {
            var request = (HttpWebRequest)WebRequest.Create(url);
            //request.Timeout = 10000; //10秒
            var response = (HttpWebResponse)request.GetResponse();
            var responseString = new StreamReader(response.GetResponseStream()).ReadToEnd();
            return responseString;
        }

        public static string PostHttpData(string url, string data)
        {
            var request = (HttpWebRequest)WebRequest.Create(url);
            var buffer = Encoding.UTF8.GetBytes(data);
            request.Method = "POST";
            request.ContentType = "application/json";
            request.ContentLength = buffer.Length;
            var strm = request.GetRequestStream();
            strm.Write(buffer, 0, buffer.Length);
            strm.Flush();
            var response = (HttpWebResponse)request.GetResponse();
            var responseString = new StreamReader(response.GetResponseStream()).ReadToEnd();
            return responseString;
        }

        public JObject JsonStrToJObject(string str)
        {
            return (JObject)JsonConvert.DeserializeObject(str);
        }

        public static int ParseInt(string intStr, int defaultValue = 0)
        {
            int parseInt;
            if (int.TryParse(intStr, out parseInt))
                return parseInt;
            return defaultValue;
        }

        public string FormatRet(RetObj obj)
        {
            //使用自定义日期格式，如果不使用的话，默认是ISO8601格式
            IsoDateTimeConverter timeConverter = new IsoDateTimeConverter();
            //"yyyy'-'MM'-'dd' 'HH':'mm':'ss"; 固定的符号如-需要用引号包裹，yyyy等则不用
            timeConverter.DateTimeFormat = "yyyy'-'MM'-'dd' 'HH':'mm':'ss";
            var responseString = JsonConvert.SerializeObject(obj, Formatting.Indented, timeConverter); //序列化
            Context.Response.Write(responseString);
            Context.Response.End();
            return responseString;
        }
         
        public string FormatRetEx(RetObj obj)
        {
            //使用自定义日期格式，如果不使用的话，默认是ISO8601格式
            IsoDateTimeConverter timeConverter = new IsoDateTimeConverter();
            //"yyyy'-'MM'-'dd' 'HH':'mm':'ss"; 固定的符号如-需要用引号包裹，yyyy等则不用
            timeConverter.DateTimeFormat = "yyyy'-'MM'-'dd' 'HH':'mm':'ss";
            var responseString = JsonConvert.SerializeObject(obj, Formatting.Indented, timeConverter); //序列化
            return responseString;
        }

        //单条数据转对象
        public static Object CreateTabToObj(DataTable dt)
        {
            dynamic d = new ExpandoObject();
            foreach (DataColumn cl in dt.Columns)
            {
                (d as ICollection<KeyValuePair<string, object>>).Add(new KeyValuePair<string, object>(cl.ColumnName, dt.Rows[0][cl.ColumnName].ToString()));
            }
            return d;
        }
        //数据表转对象集合
        public static List<Object> CreateTabToList(DataTable dt)
        {
            List<Object> list = new List<Object>();
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                dynamic d = new ExpandoObject();
                foreach (DataColumn cl in dt.Columns)
                {
                    (d as ICollection<KeyValuePair<string, object>>).Add(new KeyValuePair<string, object>(cl.ColumnName, dt.Rows[i][cl.ColumnName].ToString()));
                }
                list.Add(d);
            }
            return list;
        }

        public static string CheckFile(string checkStr, string file, string datas)
        {
            Regex intReg = new Regex("^[0-9]+$");//正整数
            Regex phoneReg = new Regex(@"^((0\d{2,3}-\d{7,8})|(1[35874]\d{9})|(0\d{2,3}\d{7,8}))$");//手机号正则
            Regex idReg = new Regex(@"^[A-Z0-9]+$");//系统id格式正则 规则:大写英文和数字
            Regex numReg = new Regex(@"^[+-]?\d*[.]?\d*$");//数字
            Regex emailReg = new Regex(@"^[a-zA-Z0-9_-]+@[a-zA-Z0-9_-]+(\.[a-zA-Z0-9_-]+)+$");//邮箱正则
            Regex hanziReg = new Regex(@"^[\u4e00-\u9fa5]{0,}$");//汉字
            Regex wordReg = new Regex(@"^[A-Za-z0-9]+$");//英文和数字
            if (!string.IsNullOrEmpty(checkStr)) return checkStr;
            if (!string.IsNullOrEmpty(datas))
            {
                JObject data = (JObject)JsonConvert.DeserializeObject(datas);
                string txt = data["txt"].ToString();
                if (!string.IsNullOrEmpty(file))
                {
                    if (data["len"] != null && !string.IsNullOrEmpty(data["len"].ToString()))
                    {
                        if (file.Length > int.Parse(data["len"].ToString()))
                        {
                            checkStr = txt + "长度过长";
                        }
                    }
                    if (data["isphone"] != null && !string.IsNullOrEmpty(data["isphone"].ToString()) && "true" == data["isphone"].ToString())
                    {
                        if (!phoneReg.Match(file).Success)
                        {
                            checkStr = txt + "不是电话格式";
                        }
                    }
                    if (data["isnumber"] != null && !string.IsNullOrEmpty(data["isnumber"].ToString()) && "true" == data["isnumber"].ToString())
                    {
                        if (!numReg.Match(file).Success)
                        {
                            checkStr = txt + "只能是数字";
                        }
                    }
                    if (data["isint"] != null && !string.IsNullOrEmpty(data["isint"].ToString()) && "true" == data["isint"].ToString())
                    {
                        if (!intReg.Match(file).Success)
                        {
                            checkStr = txt + "只能是正整数";
                        }
                    }
                    if (data["isid"] != null && !string.IsNullOrEmpty(data["isid"].ToString()) && "true" == data["isid"].ToString())
                    {
                        if (!idReg.Match(file).Success)
                        {
                            checkStr = txt + "格式不对";
                        }
                    }
                    if (data["isemail"] != null && !string.IsNullOrEmpty(data["isemail"].ToString()) && "true" == data["isemail"].ToString())
                    {
                        if (!emailReg.Match(file).Success)
                        {
                            checkStr = txt + "格式不对";
                        }
                    }
                    if (data["ishanzi"] != null && !string.IsNullOrEmpty(data["ishanzi"].ToString()) && "true" == data["ishanzi"].ToString())
                    {
                        if (!hanziReg.Match(file).Success)
                        {
                            checkStr = txt + "只能是汉字";
                        }
                    }
                    if (data["isword"] != null && !string.IsNullOrEmpty(data["isword"].ToString()) && "true" == data["isword"].ToString())
                    {
                        if (!wordReg.Match(file).Success)
                        {
                            checkStr = txt + "只能是英文和数字";
                        }
                    }
                }
                else if (data["nonull"] != null && !string.IsNullOrEmpty(data["nonull"].ToString()) && "true" == data["nonull"].ToString())
                {
                    if (string.IsNullOrEmpty(file))
                    {
                        checkStr = txt + "不能为空";
                    }
                }
            }
            return checkStr;
        }

        //获取系统服务器时间
        [WebMethod]
        public DateTime GetSvrTime()
        {
            DateTime dt = (DateTime)SQLHelper.ExecuteScalar("select GetDate()");
            Context.Response.Write(dt.ToString());
            Context.Response.End();
            return dt;
        }

        //更新用户头像
        [WebMethod]
        public string SetUserHeadImg(string UserGUID, string AvatarUrl)
        {
            RetObj obj = new RetObj();

            string sqlUpdateHeadImg = "update Sys_UUUU set AvatarUrl = @AvatarUrl where GUID = @UserGUID";
            SqlParameter[] sp = { new SqlParameter("@AvatarUrl", AvatarUrl),
                                  new SqlParameter("@UserGUID", UserGUID)
            };
            SQLHelper.ExecuteNonQuery(sqlUpdateHeadImg, sp);

            obj.errNo = "0";
            obj.errDesc = "";

            return FormatRet(obj);
        }

        //根据UserID或OpenID来获取UserInfo
        [WebMethod]
        public string GetUserInfo(string UserID_Or_OpenID)
        {
            RetObj obj = new RetObj();

            if (SQLInjection.IsAttack(UserID_Or_OpenID))
            {
                obj.errNo = "101";
                obj.errDesc = "关键字里不能含有特殊字符";
                return FormatRet(obj);
            }

            if (UserID_Or_OpenID != "")
            {
                string sqlstr = "EXEC P_SE_GetUserInfo_Mobile @UserID_Or_OpenID";
                SqlParameter[] sp = { new SqlParameter("@UserID_Or_OpenID", UserID_Or_OpenID) };
                DataTable dt = SQLHelper.ExecuteDataTable(sqlstr, sp);

                if (dt.Rows.Count == 0)
                {
                    obj.errNo = "102";
                    obj.errDesc = "用户不存在或已离职！";
                }
                else
                {
                    UserInfo usi = new UserInfo();
                    usi.UserGUID = dt.Rows[0]["UserGUID"].ToString();
                    usi.UserName = dt.Rows[0]["UserName"].ToString();
                    usi.UserID = dt.Rows[0]["UserID"].ToString();
                    usi.MobileNum = dt.Rows[0]["MobileNum"].ToString();
                    usi.AvatarUrl = dt.Rows[0]["AvatarUrl"].ToString();
                    usi.Power = dt.Rows[0]["XSPower"].ToString();
                    usi.GXKLPower = dt.Rows[0]["GXKLPower"].ToString();
                    usi.LSKLPower = dt.Rows[0]["LSKLPower"].ToString();

                    obj.errNo = "0";
                    obj.errDesc = "";
                    obj.retVal = usi;
                }
            }
            else
            {
                obj.errNo = "103";
                obj.errDesc = "UserID不能为空！";
            }

            return FormatRet(obj);
        }

        //根据jsCode来获取UserGUID,jsCode是微信用户登录时产生的有时效验证码
        [WebMethod]
        public string GetOpenID(string jsCode)
        {
            const string cUrl = "https://api.weixin.qq.com/sns/jscode2session?appid={0}&secret={1}&js_code={2}&grant_type=authorization_code";

            string url = string.Format(cUrl, cAppID, cAppSecret, jsCode);
            string resp = PostHttpData(url);
            Context.Response.Write(resp);
            Context.Response.End();
            return resp;
        }

        //将OpenID和用户绑定，绑定时需要验证用户的UserID和Pswd
        [WebMethod]
        public string BindOpenID(string OpenID, string UserID, string Pswd, string AvatarUrl)
        {
            RetObj obj = new RetObj();

            if (OpenID != "")
            {
                if (UserID == "")  //UserID为空
                {
                    obj.errNo = "101";
                    obj.errDesc = "UserID为空，无法绑定OpenID！";
                }
                else //UserID不为空
                {
                    if (OpenID != "***" && UserID == "***")  //限定老板账号只能在老板自己的微信上登录
                    {
                        obj.errNo = "105";
                        obj.errDesc = "该ERP账号不允许绑定其他微信号！";
                        return FormatRet(obj);
                    }

                    string sqlstr2 = "EXEC P_SE_GetUserInfo_Mobile @UserID";
                    SqlParameter[] sp2 = { new SqlParameter("@UserID", UserID) };
                    DataTable dt2 = SQLHelper.ExecuteDataTable(sqlstr2, sp2);

                    if (dt2.Rows.Count == 0) //UserID不正确
                    {
                        obj.errNo = "102";
                        obj.errDesc = "UserID不正确或已离职，无法绑定OpenID！";
                    }
                    else //UserID正确
                    {
                        string psw = dt2.Rows[0]["PWAES"].ToString();
                        if (psw == "" || Pswd == "*" + UserID || Pswd == AES.AESDecrypt(psw)) //Pswd正确
                        {
                            string sqlstr3 = "update Sys_UUUU set OpenID = '' where OpenID = @OpenID";
                            SqlParameter[] sp3 = { new SqlParameter("@OpenID", OpenID)
                            };
                            SQLHelper.ExecuteNonQuery(sqlstr3, sp3);

                            string sqlstr4 = "update Sys_UUUU set OpenID = @OpenID,AvatarUrl = @AvatarUrl where UserID = @UserID";
                            SqlParameter[] sp4 = { new SqlParameter("@OpenID", OpenID),
                                                    new SqlParameter("@AvatarUrl", AvatarUrl),
                                                    new SqlParameter("@UserID", UserID)
                            };
                            SQLHelper.ExecuteNonQuery(sqlstr4, sp4);

                            UserInfo usi = new UserInfo();
                            usi.UserGUID = dt2.Rows[0]["UserGUID"].ToString();
                            usi.UserName = dt2.Rows[0]["UserName"].ToString();
                            usi.UserID = dt2.Rows[0]["UserID"].ToString();
                            usi.MobileNum = dt2.Rows[0]["MobileNum"].ToString();
                            usi.AvatarUrl = AvatarUrl; //绑定时这个数据还没有保存到数据库
                            usi.Power = dt2.Rows[0]["XSPower"].ToString();
                            usi.GXKLPower = dt2.Rows[0]["GXKLPower"].ToString();
                            usi.LSKLPower = dt2.Rows[0]["LSKLPower"].ToString();

                            obj.errNo = "0";
                            obj.errDesc = "绑定OpenID成功！";
                            obj.retVal = usi;
                        }
                        else //Pswd不正确
                        {
                            obj.errNo = "103";
                            obj.errDesc = "Pswd不正确，无法绑定OpenID！";
                        }
                    }
                }
            }
            else
            {
                obj.errNo = "104";
                obj.errDesc = "OpenID不能为空！";
            }

            return FormatRet(obj);
        }

        //登录验证
        [WebMethod]
        public string LoginVerifies(string OpenID)
        {
            RetObj obj = new RetObj();

            if (OpenID != "")
            {
                string sqlstr = "EXEC P_SE_GetUserInfo_Mobile @OpenID ";
                SqlParameter[] sp = { new SqlParameter("@OpenID", OpenID) };
                DataTable dt = SQLHelper.ExecuteDataTable(sqlstr, sp);

                if (dt.Rows.Count == 0)
                {
                    obj.errNo = "101";
                    obj.errDesc = "OpenID未绑定！";
                }
                else
                {
                    UserInfo usi = new UserInfo();
                    usi.UserGUID = dt.Rows[0]["UserGUID"].ToString();
                    usi.UserName = dt.Rows[0]["UserName"].ToString();
                    usi.UserID = dt.Rows[0]["UserID"].ToString();
                    usi.MobileNum = dt.Rows[0]["MobileNum"].ToString();
                    usi.AvatarUrl = dt.Rows[0]["AvatarUrl"].ToString();
                    usi.Power = dt.Rows[0]["XSPower"].ToString();
                    usi.GXKLPower = dt.Rows[0]["GXKLPower"].ToString();
                    usi.LSKLPower = dt.Rows[0]["LSKLPower"].ToString();

                    obj.errNo = "0";
                    obj.errDesc = "";
                    obj.retVal = usi;
                }
            }
            else
            {
                obj.errNo = "103";
                obj.errDesc = "OpenID不能为空！";
            }

            return FormatRet(obj);
        }

        //获取用户待办、已办、发起、传阅的流程
        //State取值范围：全部，已读，未读
        //TaskType取值范围：Wait待办,Complete已办,SelfCreate发起,MakeCopy传阅
        //IdxPage从1开始
        //ViewClass取值范围：FlowClass按流程类型、DeptClass按部门、ProjectClass按项目
        //ClassGUID根据ViewClass不同对应不同的GUID
        [WebMethod]
        public string GetTaskList(string UserGUID, string TaskType, string State, string KeyWord, int TimeArea, string ViewClass,
            string ClassGUID, string ParentGUID, int IdxPage, int PageCount)
        {
            RetObj obj = new RetObj();

            int RowNumBegin = (IdxPage - 1) * PageCount;
            int RowNumEnd = IdxPage * PageCount;

            try
            {
                string sqlstr = "EXEC P_SE_GetTaskList_Mobile @UserGUID,@TaskType,@State,@KeyWord,@TimeArea,@ViewClass,@ClassGUID,@ParentGUID,@RowNumBegin,@RowNumEnd";
                SqlParameter[] sp = { new SqlParameter("@UserGUID", UserGUID),
                                      new SqlParameter("@TaskType", TaskType),
                                      new SqlParameter("@State", State),
                                      new SqlParameter("@KeyWord", KeyWord),
                                      new SqlParameter("@TimeArea", TimeArea),
                                      new SqlParameter("@ViewClass", ViewClass),
                                      new SqlParameter("@ClassGUID", ClassGUID),
                                      new SqlParameter("@ParentGUID", ParentGUID),
                                      new SqlParameter("@RowNumBegin", RowNumBegin),
                                      new SqlParameter("@RowNumEnd", RowNumEnd)
                };
                DataTable dt = SQLHelper.ExecuteDataTable(sqlstr, sp);

                obj.errNo = "0";
                obj.errDesc = "";
                obj.retVal = dt;
            }
            catch (Exception E)
            {
                throw new Exception(E.Message);
            }

            return FormatRet(obj);
        }

        //获取流程数据
        [WebMethod]
        public string GetFlowInfo(string FlowGUID, string UserGUID)
        {
            RetObj obj = new RetObj();
            FlowInfo flow = new FlowInfo();

            try
            {
                string sqlstr = "EXEC P_SE_GetFlowMData_Mobile @FlowGUID";
                SqlParameter[] sp = { new SqlParameter("@FlowGUID", FlowGUID) };
                DataTable dt = SQLHelper.ExecuteDataTable(sqlstr, sp);

                flow.formData = new List<FormData>();
                if (dt.Rows.Count > 0)  //流程主表数据查询出来只有一行
                {
                    for (int n = 0; n < dt.Columns.Count; n++)
                    {
                        FormData fd = new FormData();
                        fd.Name = dt.Columns[n].Caption;
                        fd.Value = dt.Rows[0][n].ToString();
                        flow.formData.Add(fd);
                    }
                }

                string sqlstr2 = "EXEC P_SE_GetFlowDData_Mobile @FlowGUID";
                SqlParameter[] sp2 = { new SqlParameter("@FlowGUID", FlowGUID) };
                DataTable dt2 = SQLHelper.ExecuteDataTable(sqlstr2, sp2);
                //flow.tableData = dt2;  //直接返回给客户端可以减少服务器计算压力，但是会增加传输压力，暂时用服务器计算的方式

                int priorOrder = 0;
                int nowOrder = 0;
                int nextOrder = 0;
                flow.tableData = new List<TableData>();
                string title = "";
                JArray ja = new JArray();
                if (dt2.Rows.Count > 0)  //循环解析流程子表数据
                {
                    for (int r = 0; r < dt2.Rows.Count; r++)
                    {
                        nowOrder = dt2.Rows[r].Field<int>("表格顺序");
                        if (r + 1 < dt2.Rows.Count)
                        {
                            nextOrder = dt2.Rows[r + 1].Field<int>("表格顺序");
                        }
                        else
                        {
                            nextOrder = 0;
                        }
                        if (priorOrder != nowOrder)
                        {
                            title = dt2.Rows[r].Field<string>("表格名称");
                            ja.Clear();
                            priorOrder = nowOrder;
                        }

                        JObject jo = new JObject();
                        for (int c = 0; c < dt2.Columns.Count; c++)
                        {
                            object data = dt2.Rows[r][c];
                            if (dt2.Columns[c].ColumnName == "表格名称" || dt2.Columns[c].ColumnName == "表格顺序") { continue; }
                            if (data != System.DBNull.Value)
                            {
                                if (dt2.Columns[c].DataType.ToString() == "System.DateTime")
                                {
                                    DateTime d = (DateTime)data;
                                    jo.Add(dt2.Columns[c].ColumnName, d.ToString("yyyy-MM-dd"));
                                }
                                else
                                {
                                    jo.Add(dt2.Columns[c].ColumnName, data.ToString());
                                }
                            }
                        }
                        ja.Add(jo);

                        if (nextOrder != nowOrder)
                        {
                            TableData td = new TableData();
                            td.Title = title;
                            td.Table = new JArray(ja);
                            flow.tableData.Add(td);
                        }
                    }
                }

                string sqlstr3 = "EXEC P_SE_GetFlowAnnex_Mobile @FlowGUID";
                SqlParameter[] sp3 = { new SqlParameter("@FlowGUID", FlowGUID) };
                DataTable dt3 = SQLHelper.ExecuteDataTable(sqlstr3, sp3);
                flow.annexData = DataTableToModel.ToListModel<AnnexData>(dt3);

                string sqlstr4 = "EXEC P_SE_GetRelationFlow_Mobile @FlowGUID ";
                SqlParameter[] sp4 = { new SqlParameter("@FlowGUID", FlowGUID) };
                DataTable dt4 = SQLHelper.ExecuteDataTable(sqlstr4, sp4);
                flow.relationData = DataTableToModel.ToListModel<RelationData>(dt4);

                string sqlstr5 = "EXEC P_SE_GetFlowOpinion_Mobile @FlowGUID ";
                SqlParameter[] sp5 = { new SqlParameter("@FlowGUID", FlowGUID) };
                DataTable dt5 = SQLHelper.ExecuteDataTable(sqlstr5, sp5);
                flow.opinionData = DataTableToModel.ToListModel<OpinionData>(dt5);

                string sqlstr6 = "select TaskType from WF_FlowNodePassState where StateNum in (0,1) and MGUID = @FlowGUID and DoUserGUID = @UserGUID";
                SqlParameter[] sp6 = { new SqlParameter("@FlowGUID", FlowGUID),
                                       new SqlParameter("@UserGUID", UserGUID)
                };
                DataTable dt6 = SQLHelper.ExecuteDataTable(sqlstr6, sp6);

                string sqlstr7 = "select a.IsAllowBack,a.IsCommunicate,a.NodeName from WF_FlowNode a inner join WF_WorkflowContentM b "
                                + "  on b.WorkflowGUID = a.WorkflowGUID and a.GUID = b.NonceNodeGUID "
                                + "where b.GUID = @FlowGUID ";
                SqlParameter[] sp7 = { new SqlParameter("@FlowGUID", FlowGUID) };
                DataTable dt7 = SQLHelper.ExecuteDataTable(sqlstr7, sp7);
                flow.indexNodeName = dt7.Rows[0]["NodeName"].ToString();

                flow.isAllowRebut = false;
                flow.isAllowConsult = false;
                if (dt6.Rows.Count > 0)
                {
                    if (dt6.Rows[0]["TaskType"].ToString() == "A")
                    {
                        if (dt7.Rows.Count > 0)
                        {
                            if (dt7.Rows[0]["IsAllowBack"].ToString() == "True")
                            {
                                flow.isAllowRebut = true;
                            }
                            if (dt7.Rows[0]["IsCommunicate"].ToString() == "True")
                            {
                                flow.isAllowConsult = true;
                            }
                        }
                    }
                }

                obj.errNo = "0";
                obj.errDesc = "";
                obj.retVal = flow;
            }
            catch (Exception E)
            {
                throw new Exception(E.Message);
            }

            return FormatRet(obj);
        }

        //流程提交
        [WebMethod]
        public string FlowSubmit(string FlowGUID, string DoUserGUID, string Agree, string Opinion)
        {
            RetObj obj = new RetObj();

            if (FlowGUID == "" || DoUserGUID == "")
            {
                obj.errNo = "101";
                obj.errDesc = "流程GUID、操作人GUID不能为空！";

                return FormatRet(obj);
            }

            try
            {
                Opinion = Opinion + "【手机审批】";

                string sqlPassState =
                    "SELECT TOP 1 ps.GUID,ps.MGUID,ps.NodeGUID,nd.NodeName,ps.TaskType,ps.TransmitUserGUID,us.UserName " +
                    "FROM WF_FlowNodePassState ps " +
                    "    inner join WF_FlowNode nd on nd.GUID = ps.NodeGUID " +
                    "    inner join HR_UserArchives us on us.GUID = ps.DoUserGUID " +
                    "WHERE MGUID = @FlowGUID AND DoUserGUID = @DoUserGUID AND ISNULL(StateNum,'')<> '2' ";
                SqlParameter[] sp = { new SqlParameter("@FlowGUID", FlowGUID),
                                      new SqlParameter("@DoUserGUID", DoUserGUID)
                };
                DataTable dt = SQLHelper.ExecuteDataTable(sqlPassState, sp);

                string OpinionGUID = "";

                if (dt.Rows.Count > 0) //必须存在流转记录才能提交
                {
                    string sqlIsHaveOpinion = "SELECT GUID FROM dbo.WF_FlowNodePassOpinion WHERE DoManGUID = @DoUserGUID "
                        + "AND DoType='10' AND MGUID = @FlowGUID";
                    SqlParameter[] sp2 = { new SqlParameter("@FlowGUID", FlowGUID),
                                           new SqlParameter("@DoUserGUID", DoUserGUID)
                    };
                    DataTable dt2 = SQLHelper.ExecuteDataTable(sqlIsHaveOpinion, sp2);

                    if (dt2.Rows.Count > 0)  //已有意见记录则更新意见
                    {
                        OpinionGUID = dt2.Rows[0].Field<string>("GUID");
                        string sqlUpdateOpinion = "Update WF_FlowNodePassOpinion set DoTime=GETDATE(),ApprovalResults = @Agree, "
                            + "ApprovalOpinion = @Opinion where GUID = @FlowGUID";
                        SqlParameter[] sp3 = { new SqlParameter("@Agree", Agree),
                                               new SqlParameter("@Opinion", Opinion),
                                               new SqlParameter("@FlowGUID", FlowGUID)
                        };
                        SQLHelper.ExecuteNonQuery(sqlUpdateOpinion, sp3);
                    }
                    else //无意见记录则新增意见
                    {
                        OpinionGUID = SQLHelper.GetGUID();

                        string sqlInsertOpinion = "insert into dbo.WF_FlowNodePassOpinion(GUID,MGUID,NodeGUID,NodeName, "
                            + "DoManGUID,DoManName,ApprovalResults,ApprovalOpinion,DoType,DoTypeName,DoTime,InternalMID,GetMan) "
                            + "VALUES (@GUID,@MGUID,@NodeGUID,@NodeName,@DoManGUID,@DoManName,@ApprovalResults,@ApprovalOpinion, "
                            + "@DoType,@DoTypeName,@DoTime,@InternalMID,@GetMan) ";
                        SqlParameter[] sp4 = { new SqlParameter("@GUID", OpinionGUID),
                                           new SqlParameter("@MGUID", FlowGUID),
                                           new SqlParameter("@NodeGUID", dt.Rows[0].Field<string>("NodeGUID")),
                                           new SqlParameter("@NodeName", dt.Rows[0].Field<string>("NodeName")),
                                           new SqlParameter("@DoManGUID", DoUserGUID),
                                           new SqlParameter("@DoManName", dt.Rows[0].Field<string>("UserName")),
                                           new SqlParameter("@ApprovalResults", Agree),
                                           new SqlParameter("@ApprovalOpinion", Opinion),
                                           new SqlParameter("@DoType", "10"),
                                           new SqlParameter("@DoTypeName", "提交"),
                                           new SqlParameter("@DoTime", DateTime.Now.ToString()),
                                           new SqlParameter("@InternalMID", "手机审批"),
                                           new SqlParameter("@GetMan", dt.Rows[0].Field<string>("TransmitUserGUID"))
                        };
                        SQLHelper.ExecuteNonQuery(sqlInsertOpinion, sp4);
                    }

                    string sqlIsReject =
                        "SELECT Top 1 GUID FROM dbo.WF_FlowNodePassState WHERE 1=1 " +
                        "and ISNULL(FlowDesc,'')<> '' " +
                        "and MGUID = @MGUID " +
                        "and NonceNodeGUID = @NonceNodeGUID " +
                        "and DoUserGUID = @DoUserGUID " +
                        "and GeiTime = (select MAX(GeiTime) as GetTime from WF_FlowNodePassState where 1=1 " +
                        "    and MGUID = @MGUID " +
                        "    and NonceNodeGUID = @NonceNodeGUID " +
                        "    and DoUserGUID = @DoUserGUID) ";
                    SqlParameter[] sp5 = { new SqlParameter("@MGUID", FlowGUID),
                                       new SqlParameter("@NonceNodeGUID", dt.Rows[0].Field<string>("NodeGUID")),
                                       new SqlParameter("@DoUserGUID", DoUserGUID)
                    };
                    DataTable dt5 = SQLHelper.ExecuteDataTable(sqlIsReject, sp5);

                    string strIsRejecte = "0";
                    if (dt5.Rows.Count > 0)
                    {
                        strIsRejecte = "1";
                    }

                    string strTaskType = dt.Rows[0].Field<string>("TaskType"); //任务类型

                    if (strTaskType == "A" || strTaskType == "D" || strTaskType == "E")
                    {
                        const string cUrl = "http://**.**.com/SHS/SHService.asmx/FlowGeneralTransfer?FlowGUID={0}&Grade={1}&DoUserGUID={2}&IsReject={3}";

                        string url = string.Format(cUrl, FlowGUID, "0", DoUserGUID, strIsRejecte);
                        string rsFlowSubmitRet = PostHttpData(url);

                        string sqlUptOpi1 = "UPDATE dbo.WF_FlowNodePassOpinion SET DoType = '20', DoTypeName = '提交', "
                            + "DoTime = GETDATE() WHERE ISNULL(DoType,'') <> '20' AND ISNULL(GetManName,'') NOT LIKE '%归档%' "
                            + "AND GUID = @GUID";
                        SqlParameter[] sp6 = { new SqlParameter("@GUID", OpinionGUID) };
                        SQLHelper.ExecuteNonQuery(sqlUptOpi1, sp6);

                        string sqlUptOpi2 = "UPDATE dbo.WF_FlowNodePassOpinion SET GetManName='归档' WHERE ISNULL(GetMan,'')=',' "
                            + "AND ISNULL(DoTypeName,'')='提交' AND GetManName IS NULL and GUID = @GUID";
                        SqlParameter[] sp7 = { new SqlParameter("@GUID", OpinionGUID) };
                        SQLHelper.ExecuteNonQuery(sqlUptOpi2, sp7);

                        if (rsFlowSubmitRet.IndexOf("Error") > 0)
                        {
                            string sqlDelOpi = "delete from dbo.WF_FlowNodePassOpinion where GUID = @GUID";
                            SqlParameter[] sp8 = { new SqlParameter("@GUID", OpinionGUID) };
                            SQLHelper.ExecuteNonQuery(sqlDelOpi, sp8);
                        }

                        obj.errNo = "0";
                        obj.errDesc = "";
                    }
                    else if (strTaskType == "B")
                    {
                        const string cUrl2 = "http://**.**.com/SHS/SHService.asmx/SetCopyTasksStateFix?FlowGUID={0}&DoUserGUID={1}&TaskGUID={2}";

                        string url2 = string.Format(cUrl2, FlowGUID, DoUserGUID, dt.Rows[0].Field<string>("GUID"));
                        string rsFlowSubmitRet = PostHttpData(url2);

                        if (rsFlowSubmitRet != "")
                        {
                            obj.errNo = "0";
                            obj.errDesc = "";
                        }
                        else
                        {
                            obj.errNo = "104";
                            obj.errDesc = "SetCopyTasksStateFix错误";
                        }
                    }
                    else if (strTaskType == "C")
                    {
                        string sqlUptOpi3 = "UPDATE dbo.WF_FlowNodePassOpinion SET GetManName=dbo.Fun_Get_UserNames(GetMan) "
                            + "WHERE GUID = @GUID AND ISNULL(GetMan,'')<>',' AND ISNULL(GetManName,'') NOT LIKE '%归档%'";
                        SqlParameter[] sp8 = { new SqlParameter("@GUID", OpinionGUID) };
                        SQLHelper.ExecuteNonQuery(sqlUptOpi3, sp8);

                        const string cUrl3 = "http://**.**.com/SHS/SHService.asmx/SetCommunicateStateFix?FlowGUID={0}&DoUserGUID={1}&TaskGUID={2}";

                        string url3 = string.Format(cUrl3, FlowGUID, DoUserGUID, dt.Rows[0].Field<string>("GUID"));
                        string rsFlowSubmitRet = PostHttpData(url3);

                        if (rsFlowSubmitRet != "")
                        {
                            obj.errNo = "0";
                            obj.errDesc = "";
                        }
                        else
                        {
                            obj.errNo = "105";
                            obj.errDesc = "SetCommunicateStateFix错误";
                        }
                    }
                    else
                    {
                        obj.errNo = "102";
                        obj.errDesc = "任务类型错误，请联系系统管理员！";
                    }
                }
                else
                {
                    obj.errNo = "103";
                    obj.errDesc = "流转记录不存在，请联系系统管理员！";

                }
            }
            catch (Exception E)
            {
                throw new Exception(E.Message);
            }

            obj.retVal = "";
            return FormatRet(obj);
        }

        //获取流程驳回节点列表
        [WebMethod]
        public string GetFlowRebutList(string FlowGUID, string DoUserGUID)
        {
            RetObj obj = new RetObj();

            string sqlGetRebutList =
                "select NodeGUID,NodeName,NodeDoUser from ( " +
                "    select distinct b.GUID as NodeGUID,b.NodeName + ' [' + c.UserName + ']' as NodeName, " +
                "    DoUserGUID as NodeDoUser,RecNo " +
                "    from dbo.WF_FlowNodePassState a " +
                "    inner join WF_FlowNode b on b.GUID = a.NodeGUID " +
                "    inner join HR_UserArchives c on c.GUID = a.DoUserGUID " +
                "    where TaskType = 'A' " +
                "    and ISNULL(DoUserGUID, '') <> '' " +
                "    and ISNULL(NodeGUID, '') <> '' " +
                "    and ISNULL(RecNo, '')  <> '' " +
                "    and MGUID = @FlowGUID " +
                "    and DoUserGUID <> @DoUserGUID " +
                ") t order by RecNo ";

            SqlParameter[] sp = { new SqlParameter("@FlowGUID", FlowGUID),
                                  new SqlParameter("@DoUserGUID", DoUserGUID)
            };
            DataTable dt = SQLHelper.ExecuteDataTable(sqlGetRebutList, sp);

            obj.errNo = "0";
            obj.errDesc = "";
            obj.retVal = dt;

            return FormatRet(obj);
        }

        //流程驳回
        [WebMethod]
        public string FlowRebut(string FlowGUID, string DoUserGUID, string NodeGUID, string NodeDoUser, string Opinion)
        {
            RetObj obj = new RetObj();

            if (FlowGUID == "" || DoUserGUID == "" || NodeGUID == "")
            {
                obj.errNo = "101";
                obj.errDesc = "流程GUID、操作人GUID、被驳回节点GUID不能为空！";

                return FormatRet(obj);
            }

            try
            {
                if (NodeDoUser == "")  //如果驳回节点的人员为空（正常情况不会出现）
                {
                    const string cUrl = "http://**.**.com/SHS/SHService.asmx/GetFlowNextNodeApprover?FlowGUID={0}&NodeGUID={1}";

                    string url = string.Format(cUrl, FlowGUID, NodeGUID);
                    NodeDoUser = PostHttpData(url);
                    if (NodeDoUser.IndexOf(",") > 0)
                    {
                        NodeDoUser = NodeDoUser.Substring(1, NodeDoUser.IndexOf(",") - 1);
                    }
                }

                const string cUrl2 = "http://**.**.com/SHS/SHService.asmx/FlowRejectTransfer?FlowGUID={0}&DoType={1}&Grade={2}"
                    + "&strPriorNodeGUID={3}&strExternalUserGUID={4}&strExtOpinion={5}&DoUserGUID={6}&strIsReject={7}";

                string url2 = string.Format(cUrl2, FlowGUID, "5", "0", NodeGUID, NodeDoUser, Opinion, DoUserGUID, "0");
                var request2 = (HttpWebRequest)WebRequest.Create(url2);
                var response2 = (HttpWebResponse)request2.GetResponse();
                var responseString2 = new StreamReader(response2.GetResponseStream()).ReadToEnd();

                obj.errNo = "0";
                obj.errDesc = responseString2;

                return FormatRet(obj);
            }
            catch (Exception E)
            {
                throw new Exception(E.Message);
            }
        }

        //获取用户清单
        [WebMethod]
        public string GetUserList(string KeyWord, int IdxPage, int PageCount) //IdxPage从1开始
        {
            RetObj obj = new RetObj();

            int RowNumBegin = (IdxPage - 1) * PageCount;
            int RowNumEnd = IdxPage * PageCount;

            string sqlUserList = "EXEC P_SE_GetUserList @KeyWord,@RowNumBegin,@RowNumEnd";
            SqlParameter[] sp = { new SqlParameter("@KeyWord", KeyWord),
                                  new SqlParameter("@RowNumBegin", RowNumBegin),
                                  new SqlParameter("@RowNumEnd", RowNumEnd)
            };
            DataTable dt = SQLHelper.ExecuteDataTable(sqlUserList, sp);

            obj.errNo = "0";
            obj.errDesc = "";
            obj.retVal = dt;

            return FormatRet(obj);
        }

        //流程沟通
        [WebMethod]
        public string FlowConsult(string FlowGUID, string DoUserGUID, string ConsultUserGUIDs, string Opinion)
        {
            RetObj obj = new RetObj();

            if (FlowGUID == "" || DoUserGUID == "" || ConsultUserGUIDs == "")
            {
                obj.errNo = "101";
                obj.errDesc = "流程GUID、操作人GUID、被沟通人GUID不能为空！";

                return FormatRet(obj);
            }

            const string cUrl = "http://**.**.com/SHS/SHService.asmx/FlowCommunicateCheck?FlowGUID={0}&Grade=&DoUserGUID={1}";

            string url = string.Format(cUrl, FlowGUID, ConsultUserGUIDs);
            string strCheckRet = PostHttpData(url);
            if (strCheckRet.IndexOf("True") > 0)
            {
                const string cUrl2 = "http://**.**.com/SHS/SHService.asmx/FlowCommunicateTransfer?FlowGUID={0}&Grade={1}&strExternalUserGUID={2}"
                    + "&strExtOpinion={3}&DoUserGUID={4}";

                string url2 = string.Format(cUrl2, FlowGUID, "0", ConsultUserGUIDs, Opinion, DoUserGUID);
                string strConsultRet = PostHttpData(url2);

                if (strConsultRet != "")  //不等于空表示沟通成功
                {
                    obj.errNo = "0";
                    obj.errDesc = "";

                    return FormatRet(obj);
                }
                else
                {
                    obj.errNo = "102";
                    obj.errDesc = "沟通失败，请联系管理员！";

                    return FormatRet(obj);
                }
            }
            else
            {
                obj.errNo = "102";
                obj.errDesc = strCheckRet;

                return FormatRet(obj);
            }
        }

        //获取首页新闻列表
        [WebMethod]
        public string GetPortalNews()
        {
            RetObj obj = new RetObj();
            PortalNews pn1 = new PortalNews();
            PortalNews pn2 = new PortalNews();
            PortalNews pn3 = new PortalNews();
            List<PortalNews> arrpn = new List<PortalNews>();

            pn1.Img = "https://**.**.com/WX/wximg/news/1.jpg";
            pn1.Url = "http://**.**-town.com/coporation/detail/id/54/type_shu/9.html";
            arrpn.Add(pn1);

            pn2.Img = "https://**.**.com/WX/wximg/news/2.jpg";
            pn2.Url = "http://**.**-town.com/coporation/detail/id/55/type_shu/9.html";
            arrpn.Add(pn2);

            pn3.Img = "https://**.**.com/WX/wximg/news/3.jpg";
            pn3.Url = "http://**.**-town.com/coporation/detail/id/56/type_shu/9.html";
            arrpn.Add(pn3);

            obj.errNo = "0";
            obj.errDesc = "";
            obj.retVal = arrpn;

            return FormatRet(obj);
        }

        //获取公司新闻列表
        [WebMethod]
        public string GetCommonNews(string NewsType, string KeyWord, string UserGUID, int IdxPage, int PageCount) //NewsType取值：全部、公司新闻、红头文件、通知公告、公司制度
        {
            RetObj obj = new RetObj();

            int RowNumBegin = (IdxPage - 1) * PageCount;
            int RowNumEnd = IdxPage * PageCount;

            string sqlGetCommonNews = "EXEC P_SE_GetComFW_Mobile @NewsType,@KeyWord,@UserGUID,@RowNumBegin,@RowNumEnd";
            SqlParameter[] sp = { new SqlParameter("@NewsType", NewsType),
                                  new SqlParameter("@KeyWord", KeyWord),
                                  new SqlParameter("@UserGUID", UserGUID),
                                  new SqlParameter("@RowNumBegin", RowNumBegin),
                                  new SqlParameter("@RowNumEnd", RowNumEnd)
            };
            DataTable dt = SQLHelper.ExecuteDataTable(sqlGetCommonNews, sp);

            obj.errNo = "0";
            obj.errDesc = "";
            obj.retVal = dt;

            return FormatRet(obj);
        }

        //获取公司新闻列表
        [WebMethod]
        public string GetPortalData(string UserGUID)
        {
            RetObj obj = new RetObj();
            PortalData pd = new PortalData();

            string sqlGetTaskCount = "EXEC P_SE_GetPersonalTaskCount @UserGUID";
            SqlParameter[] sp = { new SqlParameter("@UserGUID", UserGUID) };
            DataTable dt = SQLHelper.ExecuteDataTable(sqlGetTaskCount, sp);

            if (dt.Rows.Count > 0)
            {
                pd.ToDoCount = dt.Rows[0]["ToDoCount"].ToString(); //+ "个";
                pd.MakeCopyCount = dt.Rows[0]["MakeCopyCount"].ToString(); //+ "个";
            }

            try
            {
                const string cUrl = "http://**.**.com/KL/SHKLSvr.asmx/GetKLInfo?Year={0}&Month={1}";

                string url = string.Format(cUrl, DateTime.Now.Year, DateTime.Now.Month);
                RetObj ro = JsonConvert.DeserializeObject<RetObj>(PostHttpData(url));
                string jsonStr = ro.retVal.ToString();
                JArray ja = (JArray)JsonConvert.DeserializeObject(jsonStr);
                foreach (JObject jo in ja)
                {
                    double d1 = Convert.ToDouble(int.Parse(jo["SHGC_InSum"].ToString())) / 10000;
                    double d2 = Convert.ToDouble(int.Parse(jo["YTGX_InSum"].ToString())) / 10000;
                    pd.CustCount1 = d1.ToString("#0.0"); //+ "万人次";
                    pd.CustCount2 = d2.ToString("#0.0"); //+ "万人次";
                }
            }
            catch
            {
                pd.CustCount1 = "0"; //+ "万人次";
                pd.CustCount2 = "0"; //+ "万人次";
            }

            string sqlGetSaleData = "EXEC PRE_Sell_Report @ProjectGUID,@BDate,@QueryType";
            SqlParameter[] sp2 = { new SqlParameter("@ProjectGUID", "全部"),
                                   new SqlParameter("@BDate", DateTime.Now.ToString()),
                                   new SqlParameter("@QueryType", "0")
            };
            DataTable dt2 = SQLHelper.ExecuteDataTable(sqlGetSaleData, sp2);

            if (dt.Rows.Count > 0)
            {
                decimal RG = (decimal)dt2.Rows[0]["RG"];
                decimal QY = (decimal)dt2.Rows[0]["QY"];
                decimal HK = (decimal)dt2.Rows[0]["HK"];
                pd.RGAmount = RG.ToString("0.####");// + "万";
                pd.QYAmount = QY.ToString("0.####");// + "万";
                pd.HKAmount = HK.ToString("0.####");// + "万";
            }

            obj.errNo = "0";
            obj.errDesc = "";
            obj.retVal = pd;

            return FormatRet(obj);
        }

        //获取年度每月认购金额
        [WebMethod]
        public string GetSaleRGReport(string MYProjGUID, int Year, string ProdTypeCode)
        {
            RetObj obj = new RetObj();

            string sqlSaleReport = "EXEC P_SE_GetMonthRG_Mobile @MYProjGUID,@Year,@ProdTypeCode";
            SqlParameter[] sp = { new SqlParameter("@MYProjGUID", MYProjGUID),
                                  new SqlParameter("@Year", Year),
                                  new SqlParameter("@ProdTypeCode", ProdTypeCode)
            };
            DataTable dt = SQLHelper.ExecuteDataTable(sqlSaleReport, sp);

            if (dt.Rows.Count > 0)
            {
                obj.errNo = "0";
                obj.errDesc = "";
                obj.retVal = dt;
            }
            else
            {
                obj.errNo = "101";
                obj.errDesc = "无当年销售额数据";
            }

            return FormatRet(obj);
        }

        //获取年度每月签约金额
        [WebMethod]
        public string GetSaleQYReport(string MYProjGUID, int Year, string ProdTypeCode)
        {
            RetObj obj = new RetObj();

            string sqlSaleReport = "EXEC P_SE_GetMonthQY_Mobile @MYProjGUID,@Year,@ProdTypeCode";
            SqlParameter[] sp = { new SqlParameter("@MYProjGUID", MYProjGUID),
                                  new SqlParameter("@Year", Year),
                                  new SqlParameter("@ProdTypeCode", ProdTypeCode)
            };
            DataTable dt = SQLHelper.ExecuteDataTable(sqlSaleReport, sp);

            if (dt.Rows.Count > 0)
            {
                obj.errNo = "0";
                obj.errDesc = "";
                obj.retVal = dt;
            }
            else
            {
                obj.errNo = "101";
                obj.errDesc = "无当年销售额数据";
            }

            return FormatRet(obj);
        }

        //获取年度每月回款金额
        [WebMethod]
        public string GetSaleHKReport(string MYProjGUID, int Year, string ProdTypeCode)
        {
            RetObj obj = new RetObj();

            string sqlSaleReport = "EXEC P_SE_GetMonthHK_Mobile @MYProjGUID,@Year,@ProdTypeCode";
            SqlParameter[] sp = { new SqlParameter("@MYProjGUID", MYProjGUID),
                                  new SqlParameter("@Year", Year),
                                  new SqlParameter("@ProdTypeCode", ProdTypeCode)
            };
            DataTable dt = SQLHelper.ExecuteDataTable(sqlSaleReport, sp);

            if (dt.Rows.Count > 0)
            {
                obj.errNo = "0";
                obj.errDesc = "";
                obj.retVal = dt;
            }
            else
            {
                obj.errNo = "101";
                obj.errDesc = "无当年销售额数据";
            }

            return FormatRet(obj);
        }

        //获取明源业态列表
        [WebMethod]
        public string GetSaleProdTypeList(string MYProjGUID)
        {
            RetObj obj = new RetObj();

            string sqlstr = "EXEC P_SE_GetMYProdTypeList @MYProjGUID";
            SqlParameter[] sp = { new SqlParameter("@MYProjGUID", MYProjGUID) };
            DataTable dt = SQLHelper.ExecuteDataTable(sqlstr, sp);

            if (dt.Rows.Count > 0)
            {
                obj.errNo = "0";
                obj.errDesc = "";
                obj.retVal = dt;
            }
            else
            {
                obj.errNo = "101";
                obj.errDesc = "未获取到明源项目";
            }

            return FormatRet(obj);
        }

        //获取明源项目列表
        [WebMethod]
        public string GetSaleProjList()
        {
            RetObj obj = new RetObj();

            string sqlSaleReport = "EXEC P_SE_GetMYProjectList";
            SqlParameter[] sp = { };
            DataTable dt = SQLHelper.ExecuteDataTable(sqlSaleReport, sp);

            if (dt.Rows.Count > 0)
            {
                obj.errNo = "0";
                obj.errDesc = "";
                obj.retVal = dt;
            }
            else
            {
                obj.errNo = "101";
                obj.errDesc = "未获取到明源项目";
            }

            return FormatRet(obj);
        }

        //获取流程分类列表
        //@UserGUID varchar(32),                  --待办人GUID
        //@ViewClass varchar(20) = 'FlowClass',   --FlowClass按流程类型、DeptClass按部门、ProjectClass按项目
        //@TasksType varchar(20) = 'Wait',        --流程审批状态（Wait待办、Complete已办、SelfCreate发起、MakeCopy传阅）
        [WebMethod]
        public string GetFlowTypeList(string UserGUID, string ViewClass, string TasksType, int TimeArea)
        {
            RetObj obj = new RetObj();

            string sqlSaleReport = "EXEC P_SE_GetTaskViewClass @UserGUID,@ViewClass,@TasksType,@TimeArea,0,0,-1";
            SqlParameter[] sp = { new SqlParameter("@UserGUID", UserGUID),
                                  new SqlParameter("@ViewClass", ViewClass),
                                  new SqlParameter("@TasksType", TasksType),
                                  new SqlParameter("@TimeArea", TimeArea)
            };
            DataTable dt = SQLHelper.ExecuteDataTable(sqlSaleReport, sp);

            if (dt.Rows.Count > 0)
            {
                obj.errNo = "0";
                obj.errDesc = "";
                obj.retVal = dt;
            }
            else
            {
                obj.errNo = "101";
                obj.errDesc = "未获取到流程分类列表";
            }

            return FormatRet(obj);
        }

        //房源
        [WebMethod]
        public string GetHouseList(string ProjGUID, string KeyWord, int IdxPage, int PageCount)
        {
            RetObj obj = new RetObj();

            int RowNumBegin = (IdxPage - 1) * PageCount;
            int RowNumEnd = IdxPage * PageCount;

            string sqlGetHouseList = "EXEC P_SE_GetHouseList_Mobile @ProjGUID,@KeyWord,@RowNumBegin,@RowNumEnd";

            SqlParameter[] sp = {
                new SqlParameter("@ProjGUID", ProjGUID),
                new SqlParameter("@KeyWord", KeyWord),
                new SqlParameter("@RowNumBegin", RowNumBegin),
                new SqlParameter("@RowNumEnd", RowNumEnd)
            };
            DataTable dt = SQLHelper.ExecuteDataTable(sqlGetHouseList, sp);

            obj.errNo = "0";
            obj.errDesc = "";
            obj.retVal = dt;

            return FormatRet(obj);
        }

        //获取合同清单列表
        [WebMethod]
        public string GetConList(string ProjGUID, string UserGUID, string KeyWord, int IdxPage, int PageCount)
        {
            RetObj obj = new RetObj();

            int RowNumBegin = (IdxPage - 1) * PageCount;
            int RowNumEnd = IdxPage * PageCount;

            string sql = "EXEC P_SE_GetConList_Mobile @ProjGUID,@UserGUID,@KeyWord,@RowNumBegin,@RowNumEnd";

            SqlParameter[] sp = {
                new SqlParameter("@ProjGUID", ProjGUID),
                new SqlParameter("@UserGUID", UserGUID),
                new SqlParameter("@KeyWord", KeyWord),
                new SqlParameter("@RowNumBegin", RowNumBegin),
                new SqlParameter("@RowNumEnd", RowNumEnd)
            };
            DataTable dt = SQLHelper.ExecuteDataTable(sql, sp);

            obj.errNo = "0";
            obj.errDesc = "";
            obj.retVal = dt;

            return FormatRet(obj);
        }

        //获取合同列表-归属项目树形菜单
        [WebMethod]
        public string GetProjectInfoTree(string UserGUID)
        {
            RetObj obj = new RetObj();

            string sql = "EXEC P_SE_GetProjectList @UserGUID";
            SqlParameter[] sp = { new SqlParameter("@UserGUID", UserGUID) };
            DataTable dt = SQLHelper.ExecuteDataTable(sql, sp);

            obj.errNo = "0";
            obj.errDesc = "";
            obj.retVal = dt;

            return FormatRet(obj);
        }

        //获取合同详情
        [WebMethod]
        public string GetConInfo(string ConGUID)
        {
            RetObj obj = new RetObj();

            ConInfo conInfo = new ConInfo();

            try
            {
                //表单数据formData
                string sqlstr = "EXEC P_SE_GetConInfo_Mobile @ConGUID";
                SqlParameter[] sp = { new SqlParameter("@ConGUID", ConGUID) };
                DataTable dt = SQLHelper.ExecuteDataTable(sqlstr, sp);

                conInfo.formData = new List<FormData>();
                if (dt.Rows.Count > 0)  //流程主表数据查询出来只有一行
                {
                    for (int n = 0; n < dt.Columns.Count; n++)
                    {
                        FormData fd = new FormData();
                        fd.Name = dt.Columns[n].Caption;
                        fd.Value = dt.Rows[0][n].ToString();
                        conInfo.formData.Add(fd);
                    }
                }
                //合同附件数据
                string sqlstr2 = "EXEC P_SE_GetConAnnex_Mobile @ConGUID";
                SqlParameter[] sp2 = { new SqlParameter("@ConGUID", ConGUID) };
                DataTable dt2 = SQLHelper.ExecuteDataTable(sqlstr2, sp2);

                conInfo.annexData = DataTableToModel.ToListModel<AnnexData>(dt2);
                //合同流程数据
                string sqlstr3 = "EXEC P_SE_GetConRelaFlowList_Mobile @ConGUID";
                SqlParameter[] sp3 = { new SqlParameter("@ConGUID", ConGUID) };
                DataTable dt3 = SQLHelper.ExecuteDataTable(sqlstr3, sp3);
                conInfo.conFlowData = DataTableToModel.ToListModel<ConFlowData>(dt3);
                //合同补充协议数据
                string sqlstr4 = "EXEC P_SE_GetConSuplyList_Mobile @ConGUID";
                SqlParameter[] sp4 = { new SqlParameter("@ConGUID", ConGUID) };
                DataTable dt4 = SQLHelper.ExecuteDataTable(sqlstr4, sp4);
                conInfo.conSuplyData = DataTableToModel.ToListModel<ConSuplyData>(dt4);

                obj.errNo = "0";
                obj.errDesc = "";
                obj.retVal = conInfo;
            }
            catch (Exception E)
            {
                throw new Exception(E.Message);
            }

            return FormatRet(obj);
        }

        //获取通讯录
        [WebMethod]
        public string GetPhoneList(string PID, string KeyWord)
        {
            RetObj obj = new RetObj();

            List<OrgManData> omlist = new List<OrgManData>();

            string sqlstr1 = "EXEC P_SE_GetDingDingOrg @PID, @KeyWord";
            SqlParameter[] sp1 = { new SqlParameter("@PID", PID), new SqlParameter("@KeyWord", KeyWord) };
            DataTable dt1 = SQLHelper.ExecuteDataTable(sqlstr1, sp1);

            string sqlstr2 = "EXEC P_SE_GetDingDingMan @PID, @KeyWord";
            SqlParameter[] sp2 = { new SqlParameter("@PID", PID), new SqlParameter("@KeyWord", KeyWord) };
            DataTable dt2 = SQLHelper.ExecuteDataTable(sqlstr2, sp2);

            if (dt1.Rows.Count > 0)  //ORG
            {
                for (int i = 0; i < dt1.Rows.Count; i++)
                {
                    OrgManData omdata = new OrgManData();
                    omdata.Name = dt1.Rows[i]["name"].ToString();
                    omdata.ID = "ORG" + dt1.Rows[i]["OrgID"].ToString();
                    omdata.PID = "ORG" + dt1.Rows[i]["OrgPID"].ToString();
                    omdata.Type = "ORG";
                    omdata.Order = dt1.Rows[i]["order"].ToString();
                    omdata.Mobile = "";
                    omdata.Depart = "";
                    omdata.AvatarUrl = "";
                    omdata.UserID = "";
                    omdata.Gender = "";
                    omdata.isHide = "";
                    omlist.Add(omdata);
                }
            }

            if (dt2.Rows.Count > 0)  //MAN
            {
                for (int i = 0; i < dt2.Rows.Count; i++)
                {
                    if ("4" == dt2.Rows[i]["UserStatus"].ToString()) { continue; }

                    OrgManData omdata = new OrgManData();
                    omdata.Name = dt2.Rows[i]["name"].ToString();
                    omdata.ID = "MAN" + dt2.Rows[i]["ManID"].ToString();
                    omdata.PID = "MAN" + dt2.Rows[i]["OrgID"].ToString();
                    omdata.Type = "MAN";
                    omdata.Order = dt2.Rows[i]["orderInDepts"].ToString();
                    omdata.Mobile = dt2.Rows[i]["mobile"].ToString();
                    omdata.Depart = dt2.Rows[i]["workPlace"].ToString() + "/" + dt2.Rows[i]["position"].ToString();
                    omdata.AvatarUrl = dt2.Rows[i]["AvatarUrl"].ToString();
                    omdata.UserID = dt2.Rows[i]["UserID"].ToString();
                    omdata.Gender = dt2.Rows[i]["gender"].ToString();
                    omdata.isHide = dt2.Rows[i]["isHide"].ToString();
                    omlist.Add(omdata);
                }
            }

            obj.errNo = "0";
            obj.errDesc = "";
            obj.retVal = omlist;

            return FormatRet(obj);
        }

        //填写候选人档案
        [WebMethod]
        public string PutCandidateInfo(string ManName, string ManTel, string LastCompany, string LastStation,
            string LastSalary, string WantSalary, string HighestDegree, string School, string Speciality,
            string IDCard, string MaritalStatus, string Email, string Hometown, string HomeArea, string HomeRoad,
            string ReadyTime, string CanAllopatry, string KnowWay, string OtherWay,
            DateTime BeginTime1, DateTime EndTime1, string Company1, string Station1, 
            string Salary1, string Certifier1, string CertifierTel1,
            DateTime BeginTime2, DateTime EndTime2, string Company2, string Station2, 
            string Salary2, string Certifier2, string CertifierTel2)
        {
            RetObj obj = new RetObj();

            if (ManName == "" || ManTel == "")
            {
                obj.errNo = "1";
                obj.errDesc = "姓名和电话不能为空！";
                return FormatRetEx(obj);
            }
            if (ManTel.Length != 11)
            {
                obj.errNo = "2";
                obj.errDesc = "请正确填写电话号码！";
                return FormatRetEx(obj);
            }

            string sqlstr = "EXEC P_EX_WriteCand @ManName,@ManTel,@LastCompany,@LastStation,@LastSalary,@WantSalary,"
                + "@HighestDegree,@School,@Speciality,@IDCard,@MaritalStatus,@Email,@Hometown,@HomeArea,@HomeRoad,@ReadyTime,"
                + "@CanAllopatry,@KnowWay,@OtherWay,"
                + "@BeginTime1,@EndTime1,@Company1,@Station1,@Salary1,@Certifier1,@CertifierTel1,"
                + "@BeginTime2,@EndTime2,@Company2,@Station2,@Salary2,@Certifier2,@CertifierTel2";
            SqlParameter[] sp = {   new SqlParameter("@ManName", ManName),
                                    new SqlParameter("@ManTel", ManTel),
                                    new SqlParameter("@LastCompany", LastCompany),
                                    new SqlParameter("@LastStation", LastStation),
                                    new SqlParameter("@LastSalary", LastSalary),
                                    new SqlParameter("@WantSalary", WantSalary),
                                    new SqlParameter("@HighestDegree", HighestDegree),
                                    new SqlParameter("@School", School),
                                    new SqlParameter("@Speciality", Speciality),
                                    new SqlParameter("@IDCard", IDCard),
                                    new SqlParameter("@MaritalStatus", MaritalStatus),
                                    new SqlParameter("@Email", Email),
                                    new SqlParameter("@Hometown", Hometown),
                                    new SqlParameter("@HomeArea", HomeArea),
                                    new SqlParameter("@HomeRoad", HomeRoad),
                                    new SqlParameter("@ReadyTime", ReadyTime),
                                    new SqlParameter("@CanAllopatry", CanAllopatry),
                                    new SqlParameter("@KnowWay", KnowWay),
                                    new SqlParameter("@OtherWay", OtherWay),
                                    new SqlParameter("@BeginTime1", BeginTime1),
                                    new SqlParameter("@EndTime1", EndTime1),
                                    new SqlParameter("@Company1", Company1),
                                    new SqlParameter("@Station1", Station1),
                                    new SqlParameter("@Salary1", Salary1),
                                    new SqlParameter("@Certifier1", Certifier1),
                                    new SqlParameter("@CertifierTel1", CertifierTel1),
                                    new SqlParameter("@BeginTime2", BeginTime2),
                                    new SqlParameter("@EndTime2", EndTime2),
                                    new SqlParameter("@Company2", Company2),
                                    new SqlParameter("@Station2", Station2),
                                    new SqlParameter("@Salary2", Salary2),
                                    new SqlParameter("@Certifier2", Certifier2),
                                    new SqlParameter("@CertifierTel2", CertifierTel2)
            };

            try
            {
                int nEffRow = SQLHelper.ExecuteNonQuery(sqlstr, sp);

                obj.errNo = "0";
                obj.errDesc = "登记成功！谢谢您的支持！";
                return FormatRetEx(obj);
            }
            catch(Exception e)
            {
                obj.errNo = "3";
                obj.errDesc = "登记失败，请联系HR进行人工登记！";
                obj.retVal = e.Message;
                return FormatRetEx(obj);
            }
        }
        //修改form_id
        [WebMethod]
        public string UpdateFormIdByOpenid(string form_id, string guid)
        {

            RetObj obj = new RetObj();
            if (string.IsNullOrEmpty(form_id) || string.IsNullOrEmpty(guid))
            {
                obj.errNo = "1";
                obj.errDesc = "参数不符合规范!";
                obj.retVal = "";
                return FormatRet(obj);
            }

            string sqlstr = "update Sys_UUUU set FormID=@FormID,FormTime=@FormTime where GUID =@GUID";
            DateTime nFormTime = DateTime.Now.AddDays(7);//存入过期时间
            SqlParameter[] sp = { new SqlParameter("@FormID", form_id),
                                  new SqlParameter("@FormTime", nFormTime.ToString()),
                                  new SqlParameter("@GUID", guid)
            };

            try
            {
                int nEffRow = SQLHelper.ExecuteNonQuery(sqlstr, sp);

                if (nEffRow < 0)
                {
                    obj.errNo = "1";
                    obj.errDesc = "无影响条数,修改form_id";
                    obj.retVal = "";
                }
                else
                {
                    obj.errNo = "0";
                    obj.errDesc = "修改form_id成功！";
                    obj.retVal = "";
                }
            }
            catch(Exception em)
            {
                obj.errNo = "2";
                obj.errDesc = em.Message;
                obj.retVal = 0;
            }

            return FormatRet(obj);
        }
        //拉取用户个人设置 小程序模板消息发送开关
        //IsPostMsg : 0||1
        [WebMethod]
        public string getIsPostMsgByOpenid(string guid)
        {

            RetObj obj = new RetObj();
            if (string.IsNullOrEmpty(guid))
            {
                obj.errNo = "1";
                obj.errDesc = "参数不符合规范!";
                obj.retVal = "";
                return FormatRet(obj);
            }

            string sqlstr = "select IsPostMsg from Sys_UUUU where GUID =@GUID";
            SqlParameter[] sp = { new SqlParameter("@GUID", guid)
            };

            try
            {
                DataTable dt = SQLHelper.ExecuteDataTable(sqlstr, sp);
                if (dt.Rows.Count == 0)
                {
                    obj.errNo = "102";
                    obj.errDesc = "用户不存在或已离职！";
                }
                else
                {
                    obj.errNo = "0";
                    obj.errDesc = "";
                    string retVal = "0";
                    if ("True"==dt.Rows[0]["IsPostMsg"].ToString())
                    {//这里返回的是bool
                        retVal = "1";
                    } else {
                        retVal = "0";
                    }
                    obj.retVal = retVal;
                }
            }
            catch (Exception em)
            {
                obj.errNo = "2";
                obj.errDesc = em.Message;
                obj.retVal = 0;
            }

            return FormatRet(obj);
        }

        //设置小程序模板消息发送开关
        //IsPostMsg : 0||1
        [WebMethod]
        public string UpdateIsPostMsgByOpenid(string IsPostMsg, string guid)
        {

            RetObj obj = new RetObj();
            if (string.IsNullOrEmpty(IsPostMsg) || string.IsNullOrEmpty(guid))
            {
                obj.errNo = "1";
                obj.errDesc = "参数不符合规范!";
                obj.retVal = "";
                return FormatRet(obj);
            }

            string sqlstr = "update Sys_UUUU set IsPostMsg=@IsPostMsg where GUID =@GUID";
            SqlParameter[] sp = { new SqlParameter("@IsPostMsg", IsPostMsg),
                                  new SqlParameter("@GUID", guid)
            };

            try
            {
                int nEffRow = SQLHelper.ExecuteNonQuery(sqlstr, sp);

                if (nEffRow < 0)
                {
                    obj.errNo = "1";
                    obj.errDesc = "无影响条数,修改IsPostMsg";
                    obj.retVal = "";
                }
                else
                {
                    obj.errNo = "0";
                    obj.errDesc = "修改IsPostMsg成功！";
                    obj.retVal = "";
                }
            }
            catch (Exception em)
            {
                obj.errNo = "2";
                obj.errDesc = em.Message;
                obj.retVal = 0;
            }

            return FormatRet(obj);
        }


        //获取年月周的销售报表
        [WebMethod]
        public string GetSaleReport(string ProjGUID, string ProdTypeCode ,DateTime BeginTime, DateTime EndTime)
        {

            RetObj obj = new RetObj();

            string sqlstr = "EXEC P_SE_GetSaleTotalData_Mobile @ProjGUID,@ProdTypeCode,@BeginTime,@EndTime";
            SqlParameter[] sp = { new SqlParameter("@ProjGUID", ProjGUID),
                                  new SqlParameter("@ProdTypeCode", ProdTypeCode),
                                  new SqlParameter("@BeginTime", BeginTime),
                                  new SqlParameter("@EndTime", EndTime)
            };

            try
            {
                DataTable dt = SQLHelper.ExecuteDataTable(sqlstr, sp);

                obj.errNo = "0";
                obj.errDesc = "获取成功！";
                obj.retVal = dt;
            }
            catch (Exception em)
            {
                obj.errNo = "1";
                obj.errDesc = em.Message;
                obj.retVal = 0;
            }

            return FormatRet(obj);
        }

        //获取需要禁用邮箱的人员
        [WebMethod]
        public string GetInvalidEmailMan()
        {
            RetObj obj = new RetObj();
            string sqlstr = "EXEC P_SE_GetManForInvalidEmail";
            SqlParameter[] sp = { };
            try
            {
                DataTable dt = SQLHelper.ExecuteDataTable(sqlstr, sp);
                obj.errNo = "0";
                obj.errDesc = "获取成功！";
                obj.retVal = dt;
            }
            catch (Exception em)
            {
                obj.errNo = "1";
                obj.errDesc = em.Message;
                obj.retVal = 0;
            }
            return FormatRet(obj);
        }

        [WebMethod]
        public string GetUserDay(string GUID) ///获取员工年假
        {
            RetObj obj = new RetObj();
            string Sqlstr = "select GUID,UserName,yearholiday,YearHoliHour from HR_UserArchives where guid =@GUID";
            SqlParameter[] sp = { new SqlParameter("@GUID", GUID) };
            DataTable dt = SQLHelper.ExecuteDataTable(Sqlstr, sp);
            if (dt.Rows.Count > 0)
            {
                obj.errNo = "0";
                obj.retVal = CreateTabToObj(dt);

            }
            else
            {
                obj.errNo = "1";
                obj.errDesc = "数据查询为空";
                obj.retVal = "";
            }
            return FormatRet(obj);
        } 

        [WebMethod]
        public string SaveLeaveFlow(string jsonStr) ///请假流程提交
        {
            RetObj obj = new RetObj();
            JObject jo = (JObject)JsonConvert.DeserializeObject(jsonStr);
            string Subject = jo["Subject"].ToString();
            string project = jo["ProjectGUID"].ToString();
            string UserGUID = jo["UserGUID"].ToString();
            string FlowGUID = Guid.NewGuid().ToString("N").ToUpper();
            string text2 = jo["OrgGUID"].ToString();
            string text11 = jo["text11"].ToString();
            if (text11 == "False")
            {
                text11 = "否";

            } else
            {
                text11 = "是";
            };
            string amt1 = jo["amt1"].ToString();
            string memo = jo["memo"].ToString();  //请假事由
            string Isday = "否";
            JArray ja = (JArray)jo["leavList"]; ///插入从表数据
            foreach (JToken jt in ja)
            {
                string Text1 = jt["text1"].ToString(); //请假类型
                if (string.IsNullOrEmpty(Text1))
                {
                    obj.errDesc = "请填写请假类型";
                    obj.errNo = "2";
                    return FormatRet(obj);
                }
                else
                {
                    if (Text1 == "病假")
                        Isday = "是";

                }
            }    
            string Sqlstr = "exec  P_SE_GetAskLeaveFlowMobile  @GUID,@Subject,@ProjectGUID,@UserGUID,@text2,@text11,@amt1,@memo"; //插入流程主表
            string Wdguid = Guid.NewGuid().ToString("N").ToUpper();
            SqlParameter[] sp = { new SqlParameter("@GUID", FlowGUID),
                                new SqlParameter("@Subject", Subject),
                                new SqlParameter("@ProjectGUID", project),
                                new SqlParameter("@UserGUID", UserGUID),
                                new SqlParameter("@text2", text2),
                                new SqlParameter("@text11", text11),
                                new SqlParameter("@amt1", amt1),
                                new SqlParameter("@memo", memo)};
            int nEffRow = 0;
            try
            {
                nEffRow = SQLHelper.ExecuteNonQuery(Sqlstr, sp);
                if (nEffRow > 0)
                {
                    foreach (JToken jt in ja)
                    {
                            Sqlstr = "insert into WF_WorkFlowContentD(GUID,MGUID,text1,text4,text5,text7,text8,text12,text13,text14) values"
                            + "(@GUID,@MGUID,@text1,@text4,@text5,@text7,@text8,@text12,@text13,@text14)";
                            SqlParameter[] spd = {new SqlParameter("@GUID", Guid.NewGuid().ToString("N").ToUpper()),
                                        new SqlParameter("@MGUID", FlowGUID),
                                        new SqlParameter("@text1", jt["text1"].ToString()),
                                        new SqlParameter("@text4", jt["text4"].ToString()),
                                        new SqlParameter("@text5", jt["text5"].ToString()),
                                        new SqlParameter("@text7", jt["text7"].ToString()),
                                        new SqlParameter("@text8", jt["text8"].ToString()),
                                        new SqlParameter("@text12", jt["text12"].ToString()),
                                        new SqlParameter("@text13", jt["text13"].ToString()),
                                        new SqlParameter("@text14", jt["text14"].ToString())};
                            nEffRow = 0;
                            nEffRow = SQLHelper.ExecuteNonQuery(Sqlstr, spd);
                    }
                    if (Isday == "是")
                    {
                        Sqlstr = "update WF_WorkFlowContentM set text10=@text10 where GUID=@GUID";
                        SqlParameter[] sp2 = {  new SqlParameter("@text10",Isday),
                                                    new SqlParameter("@GUID", FlowGUID) };
                        nEffRow = SQLHelper.ExecuteNonQuery(Sqlstr, sp2);
                    }
                }
            }
            catch (Exception E)
            {
                throw new Exception(E.Message);
                obj.errDesc = E.Message;
                obj.errNo = "1";
                return FormatRet(obj);
            }
            return FlowSubmit(FlowGUID, UserGUID, "同意", "请领导审核。");
        }

        [WebMethod]
        public string GetProject()
        {
            RetObj obj = new RetObj();
            String Sqlstr = "select GUID,Par_GUID,ProjectName from V_Project_Info";
            SqlParameter[] sp = { };
            DataTable dt = SQLHelper.ExecuteDataTable(Sqlstr, sp);
            if (dt.Rows.Count > 0)
            {
                obj.errNo = "0";
                obj.retVal = CreateTabToList(dt);
            }
            else
            {
                obj.errNo = "1";
                obj.errDesc = "项目信息为空";
                obj.retVal = "";
            }
            return FormatRet(obj);
        }

        [WebMethod]
        public string GetWorkDay(string jsonStr) ///获取天数
        {
            RetObj obj = new RetObj();
            //jsonStr= "{\"sDate\":\"2020-06-12\",\"sDateTub\":\"上午\",\"eDate\":\"2020-06-14\",\"eDateTub\":\"下午\",\"isJumpWeek\":true,\"shift\":\"8\"}";
            JObject jo = (JObject)JsonConvert.DeserializeObject(jsonStr);
            double Days = 0;
            string isJumpWeek = jo["isJumpWeek"].ToString();
            string sDate = jo["sDate"].ToString();
            string eDate = jo["eDate"].ToString();
            string sDateTub = jo["sDateTub"].ToString();
            string eDateTub = jo["eDateTub"].ToString();
            string checkStr = "";
            checkStr =CheckFile(checkStr, jsonStr, "{'txt':'参数对象','nonull':'true'}");
            checkStr =CheckFile(checkStr, sDate, "{'txt':'开始时间','nonull':'true'}");
            checkStr = CheckFile(checkStr, eDate, "{'txt':'结果时间','nonull':'true'}");
            checkStr = CheckFile(checkStr, sDateTub, "{'txt':'开始时段','nonull':'true'}");
            checkStr = CheckFile(checkStr, eDateTub, "{'txt':'结束时段','nonull':'true'}");
            if (!string.IsNullOrEmpty(checkStr))
            {
                obj.errNo = "2";
                obj.errDesc = checkStr;
                return FormatRet(obj);
            }
            if (jo["isJumpWeek"].ToString() == "False") //是否计算周末
            {
                string Sqlstr = "select * from dbo.Fun_Get_WorkDay(@sDate,@eDate,1)";
                SqlParameter[] sp2 = { new SqlParameter("@sDate", sDate),
                                       new SqlParameter("@eDate", eDate)};
                DataTable dt = SQLHelper.ExecuteDataTable(Sqlstr, sp2);
                Days = dt.Rows.Count;
                if (Days <= 0)
                {
                    obj.errDesc = "日期选择有误，请判断是否计算周末";
                    obj.errNo = "1";
                    return FormatRet(obj);
                }
            }
            else
            {
                Days = (DateTime.Parse(eDate) - DateTime.Parse(sDate)).Days+1;
            }

            if (sDateTub == "上午" && eDateTub == "上午")
            {
                Days = Days - 0.5;
            }
            else if (sDateTub == "下午" && eDateTub == "上午")
            {
                Days = Days -1;
            }
            else if (sDateTub == "下午" && eDateTub == "下午")
            {
                Days = Days - 0.5;
            }
            obj.errDesc = "";
            obj.errNo = "0";
           
            obj.retVal = (Days * 8).ToString();
            return FormatRet(obj);
        }

        [WebMethod]
        public string GetBusName() ///获取公司名称
        {
            RetObj obj = new RetObj();
            string Sqlstr = "select GUID,CodeName from Sys_Dictionary where CodeTypeGUID =@CodeTypeGUID";
            SqlParameter[] sp = { new SqlParameter("@CodeTypeGUID", "DD7A43A60AAE46F1B3CD9EF2EEB172EA") };
            DataTable dt = SQLHelper.ExecuteDataTable(Sqlstr, sp);
            if (dt.Rows.Count > 0)
            {
                obj.errNo = "0";
                obj.retVal = CreateTabToList(dt);
            }
            else
            {
                obj.errNo = "1";
                obj.errDesc = "公司名字数据查询为空";
                obj.retVal = "";
            }
            return FormatRet(obj);
        }


        //返回当月哪些天有会议用于标记
        //20200915
        [WebMethod]
        public string GetCurMouthCalendar(string CurMouth)
        {
            RetObj obj = new RetObj();

            //参数校验
            string checkStr = "";
            checkStr = CheckFile(checkStr, CurMouth, "{'txt':'当前月','nonull':'true'}");
            if (!string.IsNullOrEmpty(checkStr))
            {
                obj.errNo = "2";
                obj.errDesc = checkStr;
                return FormatRet(obj);
            }

            string Sqlstr = " select CONVERT(varchar,a.MeetDate,23) as MeetDate from [dbo].[AM_MeetApply] a where CONVERT(varchar,a.MeetDate,120) like '%"+ CurMouth + "%'";
            SqlParameter[] sp = { };
            DataTable dt = SQLHelper.ExecuteDataTable(Sqlstr, sp);
            if (dt.Rows.Count > 0)
            {
                obj.errNo = "0";
                obj.errDesc = "数据查询成功!";
                obj.retVal = CreateTabToList(dt);
            }
            else
            {
                obj.errNo = "3";
                obj.errDesc = "查询数据为空!";
                obj.retVal = "";
            }
            return FormatRet(obj);
        }

        //拉取会议日历
        //MeetDate  精确到天
        //20200915
        [WebMethod]
        public string GetMeetCalendar(string MeetDate)
        {
            RetObj obj = new RetObj();

            //参数校验
            string checkStr = "";
            checkStr = CheckFile(checkStr, MeetDate, "{'txt':'会议日期','nonull':'true'}");
            if (!string.IsNullOrEmpty(checkStr))
            {
                obj.errNo = "2";
                obj.errDesc = checkStr;
                return FormatRet(obj);
            }

            string Sqlstr = " select b.GUID,a.BeginTime,a.EndTime,b.MeetTitle,m.MtName,u.UserName from [dbo].[AM_MeetApply] a,[dbo].[AM_MeetCard] b left join [dbo].[Sys_UUUU]"
                +" u on b.LeadingMan=u.GUID,[dbo].[AM_MeetRoom] m where b.MGUID=a.GUID and m.GUID=a.MeetRoom and CONVERT(varchar,a.MeetDate,120) like '%" + MeetDate + "%'";
            SqlParameter[] sp = {  };
            DataTable dt = SQLHelper.ExecuteDataTable(Sqlstr, sp);
            if (dt.Rows.Count > 0)
            {
                obj.errNo = "0";
                obj.errDesc = "数据查询成功!";
                obj.retVal = CreateTabToList(dt);
            }
            else
            {
                obj.errNo = "3";
                obj.errDesc = "查询数据为空!";
                obj.retVal = "";
            }
            return FormatRet(obj);
        }

        //拉取会议明细(会议卡片)
        //返回选取会议的会议明细
        //20200915
        [WebMethod]
        public string GetMeetDetailByMeetGUID(string GUID)
        {
            RetObj obj = new RetObj();

            dynamic myObj = new ExpandoObject();
            //参数校验
            string checkStr = "";
            checkStr = CheckFile(checkStr, GUID, "{'txt':'会议GUID','nonull':'true'}");
            if (!string.IsNullOrEmpty(checkStr))
            {
                obj.errNo = "2";
                obj.errDesc = checkStr;
                return FormatRet(obj);
            }

            string Sqlstr = "exec [dbo].[PRE_MeetDetailByGUID] @GUID";
            SqlParameter[] sp = { new SqlParameter("@GUID", GUID) };
            DataTable dt = SQLHelper.ExecuteDataTable(Sqlstr, sp);
            if (dt.Rows.Count > 0)
            {
                //obj.errNo = "0";
                myObj = CreateTabToObj(dt);
                //循环取会议室配备
                string MGUID = myObj.MGUID; // dt.Rows[0]["MGUID"].ToString();
                string Sqlstr2 = "select * from [dbo].[AM_MeetRoomEquip] mre where MGUID= @MGUID";
                SqlParameter[] sp2 = { new SqlParameter("@MGUID", MGUID) };
                DataTable dt2 = SQLHelper.ExecuteDataTable(Sqlstr2, sp2);
                if (dt2.Rows.Count > 0)
                {
                    string EquipNameStr = "";
                    for (int i = 0; i < dt2.Rows.Count; i++)
                    {
                        EquipNameStr += string.IsNullOrEmpty(EquipNameStr)? dt2.Rows[i]["EquipName"].ToString() : ","+dt2.Rows[i]["EquipName"].ToString();
                    }
                    //myObj.Add("EquipNameStr", EquipNameStr);
                    myObj.EquipNameStr = EquipNameStr;
                }

                obj.errNo = "0";
                obj.errDesc = "查询数据成功!";
                obj.retVal = myObj;
            }
            else
            {
                obj.errNo = "3";
                obj.errDesc = "查询数据为空!";
                obj.retVal = "";
            }
            return FormatRet(obj);
        }

        //微信订阅模板消息支持接口-入库待发消息
        //uni.requestSubscribeMessage({})
        //是否需要联动用户表中的消息提醒设置
        //State  1/2    1:待发送  2:已发送
        //20200917
        [WebMethod]
        public string InsertWXSubscribeNotify(string OpenID,string ManGUID,string MeetGUID,string TemplateID, string Page, string HintType, string EarlyMin)
        {
            RetObj obj = new RetObj();
           
            //参数校验
            string checkStr = "";
            checkStr = CheckFile(checkStr, OpenID, "{'txt':'OpenID','nonull':'true'}");
            checkStr = CheckFile(checkStr, ManGUID, "{'txt':'用户GUID','nonull':'true'}");
            checkStr = CheckFile(checkStr, MeetGUID, "{'txt':'会议GUID','nonull':'true'}");
            checkStr = CheckFile(checkStr, TemplateID, "{'txt':'订阅消息模板ID','nonull':'true'}");
            checkStr = CheckFile(checkStr, Page, "{'txt':'跳转页面','nonull':'true'}");
            checkStr = CheckFile(checkStr, HintType, "{'txt':'消息类型','nonull':'true'}");
            checkStr = CheckFile(checkStr, EarlyMin, "{'txt':'提醒时间刻度','nonull':'true','isint':'true'}");
            //checkStr = CheckFile(checkStr, State, "{'txt':'提醒状态','nonull':'true'}");
            if (!string.IsNullOrEmpty(checkStr))
            {
                obj.errNo = "2";
                obj.errDesc = checkStr;
                return FormatRet(obj);
            }
            //仅限会议前通知单条会议 单用户仅通知一次 重复点击只生成一条通知 其它通知(如会议取消/更改) 可通知多次 
            string sql = @"if not exists (select GUID from [dbo].[AM_WXMsg] where OpenID=@OpenID1 and MeetGUID=@MeetGUID1) insert into [dbo].[AM_WXMsg](GUID,OpenID,ManGUID,MeetGUID,TemplateID,Page,HintType,EarlyMin,State) 
                VALUES (@GUID,@OpenID,@ManGUID,@MeetGUID,@TemplateID,@Page,@HintType,@EarlyMin,1)";
            SqlParameter[] sqlPar = {
                new SqlParameter("@OpenID1",OpenID),
                new SqlParameter("@MeetGUID1",MeetGUID),
                new SqlParameter("@GUID", SQLHelper.GetGUID()),
                new SqlParameter("@OpenID",OpenID),
                new SqlParameter("@ManGUID",ManGUID),
                new SqlParameter("@MeetGUID",MeetGUID),
                new SqlParameter("@TemplateID",TemplateID),
                new SqlParameter("@Page",Page),
                new SqlParameter("@HintType",HintType),
                new SqlParameter("@EarlyMin",EarlyMin)
            };
            int nEffRow = SQLHelper.ExecuteNonQuery(sql, sqlPar);
            if (nEffRow > 0)
            {
                obj.errNo = "0";
                obj.errDesc = "操作成功!";
                obj.retVal = "";
            }
            else
            {
                obj.errNo = "3";
                obj.errDesc = "操作失败,您已经设置过了!";
                obj.retVal = "";
            }
            return FormatRet(obj);
        }

        //小程序端领导行程初始化
        //data  精确到day
        //2021010
        [WebMethod]
        public string InitLeaderWay(string ManGUID,string date)
        {
            RetObj obj = new RetObj();

            //参数校验
            string checkStr = "";
            checkStr = CheckFile(checkStr, ManGUID, "{'txt':'请选取领导','nonull':'true'}");
            checkStr = CheckFile(checkStr, date, "{'txt':'请选取时间','nonull':'true'}");
            if (!string.IsNullOrEmpty(checkStr))
            {
                obj.errNo = "2";
                obj.errDesc = checkStr;
                return FormatRet(obj);
            }

            string sql = "select GUID,BeginTime,EndTime,CONVERT(varchar,ActDate,102) as ActDate,ManGUID,ActType,CreateDate from [dbo].[AM_ManSchedule] where ManGUID=@ManGUID and CONVERT(varchar,ActDate,120) like '%" + date + "%'";
            SqlParameter[] sqlPar = {
                new SqlParameter("@ManGUID",ManGUID)
            };
            DataTable dt2 = SQLHelper.ExecuteDataTable(sql, sqlPar);
            obj.errNo = "0";
            obj.errDesc = "操作成功!";
            obj.retVal = dt2;
            return FormatRet(obj);
        }

        //小程序端领导行程上报
        //根据上传的时间段和领导判断该条数据是否可入库
        //2021010
        [WebMethod]
        public string SaveLeaderWay(string UserGUID, string ManGUID, string BeginTime, string EndTime, string ActDate, string ActType)
        {
            RetObj obj = new RetObj();

            //参数校验
            string checkStr = "";
            checkStr = CheckFile(checkStr, UserGUID, "{'txt':'当前登陆用户','nonull':'true'}");
            checkStr = CheckFile(checkStr, ManGUID, "{'txt':'领导GUID','nonull':'true'}");
            checkStr = CheckFile(checkStr, BeginTime, "{'txt':'开始时间','nonull':'true','isint':'true'}");
            checkStr = CheckFile(checkStr, EndTime, "{'txt':'结束时间','nonull':'true','isint':'true'}");
            checkStr = CheckFile(checkStr, ActDate, "{'txt':'开始日期','nonull':'true'}");
            checkStr = CheckFile(checkStr, ActType, "{'txt':'行程事项','nonull':'true'}");
            if (!string.IsNullOrEmpty(checkStr))
            {
                obj.errNo = "2";
                obj.errDesc = checkStr;
                return FormatRet(obj);
            }
            if (int.Parse(BeginTime) > int.Parse(EndTime)) {
                obj.errNo = "2";
                obj.errDesc = "开始时间不可大于结束时间";
                return FormatRet(obj);
            }
            //输入的时间区间与数据库的时间区间相对比  对比条件(日期/开始时间/结束时间/领导用户GUID)
            string sql1 = "select * from ";
            //根据不通的模板生成不同的keyword列表
            String[] TimeScope = new string[] {
                "08:00", "08:30", "09:00", "09:30", "10:00", "10:30", "11:00", "11:30", "12:00", "12:30","13:00", "13:30",
                "14:00", "14:30", "15:00", "15:30", "16:00", "16:30", "17:00", "17:30", "18:00", "18:30", "19:00", "19:30"
            };
            ActDate = ActDate+ " "+ TimeScope[int.Parse(BeginTime) - 1];
            //仅限会议前通知单条会议 单用户仅通知一次 重复点击只生成一条通知 其它通知(如会议取消/更改) 可通知多次 
            string sql = @" insert into [dbo].[AM_ManSchedule](GUID,ManGUID,BeginTime,EndTime,ActDate,ActType) values(@GUID,@ManGUID,@BeginTime,@EndTime,@ActDate,@ActType)";
            SqlParameter[] sqlPar = {
                new SqlParameter("@GUID",SQLHelper.GetGUID()),
                new SqlParameter("@ManGUID",ManGUID),
                new SqlParameter("@BeginTime",BeginTime),
                new SqlParameter("@EndTime",EndTime),
                new SqlParameter("@ActDate",ActDate),
                new SqlParameter("@ActType",ActType)
            };
           
            int nEffRow = SQLHelper.ExecuteNonQuery(sql, sqlPar);
            if (nEffRow > 0)
            {
                obj.errNo = "0";
                obj.errDesc = "操作成功!";
                obj.retVal = "";
            }
            else
            {
                obj.errNo = "3";
                obj.errDesc = "操作失败,1212121!";
                obj.retVal = "";
            }
            return FormatRet(obj);

        }
        //
        //json格式字符串读取测试
        //[WebMethod]
        //public string TestJson()
        //{
        //    RetObj obj = new RetObj();

        //    string jsonStr = "{'errNo': '0', 'errDesc': {'UserID': '01271'},'retVal': ["
        //        + "{'UserGUID': '9D75200DBEC744EBACD2E208BB9BD900','UserName': '冯明强'},"
        //        + "{'UserGUID': '7775200DBEC744EBACD2E208BB9BD900','UserName': '雷震'}"
        //        + "]}";

        //    string printStr = "";

        //    JObject jo = (JObject)JsonConvert.DeserializeObject(jsonStr);
        //    printStr = printStr + jo["errNo"] + ";";
        //    JObject jo2 = (JObject)jo["errDesc"];
        //    printStr = printStr + jo2["UserID"] + ";";
        //    JArray ja = (JArray)jo["retVal"]; ;
        //    foreach (JToken jt in ja)
        //    {
        //        printStr = printStr + jt["UserGUID"] + ";";
        //        printStr = printStr + jt["UserName"] + ";";
        //    }

        //    obj.errNo = "0";
        //    obj.errDesc = "";
        //    obj.retVal = printStr;

        //    return FormatRet(obj);
        //}

    }
}
