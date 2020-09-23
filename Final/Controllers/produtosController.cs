using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Entity;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Net;
using System.Web;
using System.Web.Configuration;
using System.Web.Mvc;
using Final.Models;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis.TokenSeparatorHandlers;
using PagedList;

namespace Final.Controllers
{
    [Authorize]
    public class produtosController : Controller
    {
        private Contexto db = new Contexto();

        // GET: produtos
        public ActionResult Index(int ? page, string sortOrder, string searchString, string currentFilter)
        {
            ViewBag.CurrentSort = sortOrder;
            ViewBag.NomeParam = String.IsNullOrEmpty(sortOrder) ? "DescProduto" : "";
            ViewBag.CurrentFilter = searchString;


            var produt = from s in db.produtos
                           select s;

            if (!String.IsNullOrEmpty(searchString))
            {
                produt = db.produtos.Where(s => s.DescProduto.ToUpper().Contains(searchString.ToUpper()));
                                      
            }
            switch (sortOrder)
            {
                case "DescProduto":
                    produt = produt.OrderBy(s => s.DescProduto);
                    break;
                default:
                    produt = produt.OrderBy(s => s.DescProduto);
                    break;
            }

            int pageNumber =page ?? 1;
            int pageSize =  10;
            var produc = db.produtos.OrderBy(x => x.Id).ToPagedList(pageNumber,pageSize);



            return View(produc);


        }

        #region Search

        #endregion

        #region CRUD        
        // GET: produtos/Details/5
        public ActionResult Details(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            produto produto = db.produtos.Find(id);
            if (produto == null)
            {
                return HttpNotFound();
            }
            return View(produto);
        }

        // GET: produtos/Create
        public ActionResult Create()
        {
            return View();
        }

        // POST: produtos/Create
        // To protect from overposting attacks, enable the specific properties you want to bind to, for 
        // more details see https://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Create([Bind(Include = "Id,CodProud,DescProduto,IVA,ICMS1,ICMS2,CodICM,PercBaSerDST,PercBaSerd,tipo,NCM,RegTribEstadual,BaseICMS,VendaIntCred,VendaIntDeb,ICMCred,ICMDeb,MVAind,MVAatac,MVA4,MVA7,MVA12")] produto produto)
        {
            if (ModelState.IsValid)
            {
                db.produtos.Add(produto);
                db.SaveChanges();
                return RedirectToAction("Index");
            }

            return View(produto);
        }

        // GET: produtos/Edit/5
        public ActionResult Edit(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            produto produto = db.produtos.Find(id);
            if (produto == null)
            {
                return HttpNotFound();
            }
            return View(produto);
        }

        // POST: produtos/Edit/5
        // To protect from overposting attacks, enable the specific properties you want to bind to, for 
        // more details see https://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Edit([Bind(Include = "Id,CodProud,DescProduto,IVA,ICMS1,ICMS2,CodICM,PercBaSerDST,PercBaSerd,tipo,NCM,RegTribEstadual,BaseICMS,VendaIntCred,VendaIntDeb,ICMCred,ICMDeb,MVAind,MVAatac,MVA4,MVA7,MVA12")] produto produto)
        {
            if (ModelState.IsValid)
            {
                db.Entry(produto).State = EntityState.Modified;
                db.SaveChanges();
                return RedirectToAction("Index");
            }
            return View(produto);
        }

        // GET: produtos/Delete/5
        public ActionResult Delete(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            produto produto = db.produtos.Find(id);
            if (produto == null)
            {
                return HttpNotFound();
            }
            return View(produto);
        }

        // POST: produtos/Delete/5
        [HttpPost, ActionName("Delete")]
        [ValidateAntiForgeryToken]
        public ActionResult DeleteConfirmed(int id)
        {
            produto produto = db.produtos.Find(id);
            db.produtos.Remove(produto);
            db.SaveChanges();
            return RedirectToAction("Index");
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                db.Dispose();
            }
            base.Dispose(disposing);
        }
        #endregion


        #region Donwload

        public ActionResult Excel()
        {

            List<produto> accounts = db.produtos.ToList();

            using (ExcelPackage pck = new ExcelPackage())
            {
                ExcelWorksheet ws = pck.Workbook.Worksheets.Add("planilha1");
                ws.Cells["A1"].LoadFromCollection(accounts, true);
                // Load your collection "accounts"

                Byte[] fileBytes = pck.GetAsByteArray();
                Response.Clear();
                Response.Buffer = true;
                Response.AddHeader("content-disposition", "attachment;filename=Produtos.xlsx");
                // Replace filename with your custom Excel-sheet name.

                Response.Charset = "";
                Response.ContentType = "application/vnd.ms-excel";
                StringWriter sw = new StringWriter();
                Response.BinaryWrite(fileBytes);
                Response.End();


            }

            return RedirectToAction("Index");
        }



        #endregion

        #region Upload

        [HttpPost]
        public ActionResult Index(HttpPostedFileBase file)
        {
            string filePath = string.Empty;
            if (file != null)
            {
                string path = Server.MapPath("~/Uploads/");
                if (!Directory.Exists(path))
                {
                    Directory.CreateDirectory(path);
                }

                filePath = path + Path.GetFileName(file.FileName);
                string extension = Path.GetExtension(file.FileName);
                file.SaveAs(filePath);

                string conString = string.Empty;

                switch (extension)
                {
                    case ".xls": //Excel 97-03.
                        conString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + filePath + ";Extended Properties='Excel 8.0;HDR=YES'";
                        break;
                    case ".xlsx": //Excel 07 and above.
                        conString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + filePath + ";Extended Properties='Excel 8.0;HDR=YES'";
                        break;
                }

                DataTable dt = new DataTable();
                conString = string.Format(conString, filePath);

                using (OleDbConnection connExcel = new OleDbConnection(conString))
                {
                    using (OleDbCommand cmdExcel = new OleDbCommand())
                    {
                        using (OleDbDataAdapter odaExcel = new OleDbDataAdapter())
                        {
                            cmdExcel.Connection = connExcel;

                            //Get the name of First Sheet.
                            connExcel.Open();
                            DataTable dtExcelSchema;
                            dtExcelSchema = connExcel.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                            string sheetName = dtExcelSchema.Rows[0]["TABLE_NAME"].ToString();
                            connExcel.Close();

                            //Read Data from First Sheet.
                            connExcel.Open();
                            cmdExcel.CommandText = "SELECT * From [" + sheetName + "]";
                            odaExcel.SelectCommand = cmdExcel;
                            odaExcel.Fill(dt);
                            connExcel.Close();
                        }
                    }
                }

                conString = @"Data Source=INFO-01\SQLEXPRESS;Initial Catalog=Final.Models.Contexto;Integrated Security=True;Connect Timeout=30;Encrypt=False;TrustServerCertificate=False;ApplicationIntent=ReadWrite;MultiSubnetFailover=False";
                using (SqlConnection con = new SqlConnection(conString))
                {
                    using (SqlBulkCopy sqlBulkCopy = new SqlBulkCopy(con))
                    {
                        //Set the database table name.
                        sqlBulkCopy.DestinationTableName = "dbo.produtoes";

                        // Map the Excel columns with that of the database table, this is optional but good if you do
                        // 

                        sqlBulkCopy.ColumnMappings.Add("CODPROD", "CodProud");
                        sqlBulkCopy.ColumnMappings.Add("DESCRICAO", "DescProduto");
                        sqlBulkCopy.ColumnMappings.Add("IVA", "IVA");
                        sqlBulkCopy.ColumnMappings.Add("ALIQICMS1", "ICMS1");
                        sqlBulkCopy.ColumnMappings.Add("ALIQICMS2", "ICMS2");
                        sqlBulkCopy.ColumnMappings.Add("PERCBASEREDST", "PercBaSerDST");
                        sqlBulkCopy.ColumnMappings.Add("CODICM", "CodICM");
                        sqlBulkCopy.ColumnMappings.Add("tipo", "tipo");
                        sqlBulkCopy.ColumnMappings.Add("PERCBASERED", "PercBaSerd");
                        sqlBulkCopy.ColumnMappings.Add("NCM", "NCM");
                        sqlBulkCopy.ColumnMappings.Add("Regime de Tributação Estadual", "RegTribEstadual");
                        sqlBulkCopy.ColumnMappings.Add("Base ICMS", "BaseICMS");
                        sqlBulkCopy.ColumnMappings.Add("Venda Interna Credito", "VendaIntCred");
                        sqlBulkCopy.ColumnMappings.Add("Venda Intena Debito", "VendaIntDeb");
                        sqlBulkCopy.ColumnMappings.Add("AP Normal Aliq ICMS Credito", "ICMCred");
                        sqlBulkCopy.ColumnMappings.Add("Ap Normal Aliq ICMS Debito", "ICMDeb");
                        sqlBulkCopy.ColumnMappings.Add("MVA Ind", "MVAind");
                        sqlBulkCopy.ColumnMappings.Add("MVA Atac", "MVAatac");
                        sqlBulkCopy.ColumnMappings.Add("MVA 4%", "MVA4");
                        sqlBulkCopy.ColumnMappings.Add("MVA 7%", "MVA7");
                        sqlBulkCopy.ColumnMappings.Add("MVA 12%", "MVA12");

                        con.Open();
                        sqlBulkCopy.WriteToServer(dt);
                        con.Close();
                    }
                }
            }
            //if the code reach here means everthing goes fine and excel data is imported into database
            ViewBag.Success = "File Imported and excel data saved into database";

            return RedirectToAction("Index");
        }



        #endregion
    }
}
