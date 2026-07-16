using StandardMetals.BAL;
using StandardMetals.MODEL;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Web;
using ClosedXML.Excel;
using System.Web.Mvc;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace StandardMetals.WebApp.Controllers
{
    public class StudentProgressReportController : Controller
    {
        StudentProgressReportBAL SBal = new StudentProgressReportBAL();

        [HttpGet]
        public ActionResult Index()
        {
            ViewBag.Student = SBal.StudentList();

            // Empty batch list initially
            ViewBag.Batch = new List<StudentProgressReportMODEL>();

            return View();
        }

        [HttpPost]
        public ActionResult StudentProgressReport(StudentProgressReportMODEL model)
        {
            ViewBag.Student = SBal.StudentList();
            ViewBag.ShowDownload = true;

            return View("Index", model);
        }

        [HttpGet]
        public JsonResult BatchList(int studentId)
        {
            var batchList = SBal.BatchList(studentId);

            return Json(batchList, JsonRequestBehavior.AllowGet);
        }




        
        [HttpGet]
        public ActionResult DownloadReport(int StudentId, int BatchId)
        {
            DataSet ds = SBal.GetStudentProgressReport( StudentId, BatchId);

            if (ds == null || ds.Tables.Count == 0)
            {
                return Content("No Data Found");
            }

            // Get student name
            var student = SBal.StudentList().FirstOrDefault(x => x.StudentId == StudentId);

            string studentName = student != null ? student.StudentName.Replace(" ", "_") : "Student";

            using (XLWorkbook wb = new XLWorkbook())
            {
                if (ds.Tables.Count > 0) wb.Worksheets.Add(ds.Tables[0], "Assignment Report");

                if (ds.Tables.Count > 1) wb.Worksheets.Add(ds.Tables[1], "Attendance Report");

                if (ds.Tables.Count > 2) wb.Worksheets.Add(ds.Tables[2], "Mock Interview Report");

                if (ds.Tables.Count > 3) wb.Worksheets.Add(ds.Tables[3], "Test Report");

                using (MemoryStream stream = new MemoryStream())
                {
                    wb.SaveAs(stream);

                    return File(
                stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                studentName + "_StudentProgressReport.xlsx");
                }
            }
        }

    }
}