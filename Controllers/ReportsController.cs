using iTextGenPDF.Api.Models;
using iTextGenPDF.Api.Services.Interface;
using Microsoft.AspNetCore.Mvc;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading.Tasks;
using System.Security.Cryptography;
using System.Text;

namespace iTextGenPDF.Api.Controllers
{
    [ApiController]
    [Route("api/v1/[Controller]")]
    public class ReportsController : Controller
    {
        private readonly IGenPDFService _genPDFService;
        public ReportsController(IGenPDFService genPDFService)
        {
            _genPDFService = genPDFService;
        }

        [HttpPost("CreditNote")]
        public IActionResult CreateCreditNote(Refund creditNoteModel)
        {
            creditNoteModel.TypeCode = "81";
            var res = _genPDFService.CreateCreditDebitNote(creditNoteModel);
            if (res.StatusCode == 200)
            {
                return File(res.File, "application/octet-stream", $"CreditNote_{DateTime.Now.ToString("ddMMyyyyHHmmss")}.pdf");
            }
            else
            {
                return NotFound(res);
            }
        }
        [HttpPost("DebitNote")]
        public IActionResult CreateDebitNote(Refund debitNoteModel)
        {
            debitNoteModel.TypeCode = "80";
            var res = _genPDFService.CreateCreditDebitNote(debitNoteModel);
            if (res.StatusCode == 200)
            {
                return File(res.File, "application/octet-stream", $"DebitNote_{DateTime.Now.ToString("ddMMyyyyHHmmss")}.pdf");
            }
            else
            {
                return NotFound(res);
            }
        }

        [HttpPost("TaxInvoice")]
        public IActionResult CreateTaxInvoice(Payment payment)
        {
            var res = _genPDFService.CreateTaxInvoice(payment);
            if (res.StatusCode == 200)
            {
                return File(res.File, "application/octet-stream", $"TaxInvoice_{DateTime.Now.ToString("ddMMyyyyHHmmss")}.pdf");
            }
            else
            {
                return NotFound(res);
            }
        }

        [HttpPost("Truemove_CreditNote")]
        public IActionResult TrueCreditNote(Refund creditNoteModel)
        {
            var res = _genPDFService.CreateTrueMoveCDN(creditNoteModel);
            if (res.StatusCode == 200)
            {
                return File(res.File, "application/octet-stream", $"CreditNote_{DateTime.Now.ToString("ddMMyyyyHHmmss")}.pdf");
            }
            else
            {
                return NotFound(res);
            }
        }

        [HttpGet("EncryptPassword")]
        public ResponseModel EncryptPassword(string plainText)
        {
            return _genPDFService.EncrptPassword(plainText);
        }

        [HttpGet("DecryptPassword")]
        public ResponseModel DecryptPassword(string cipherText)
        {
            return _genPDFService.DecrptPassword(cipherText);
        }

    }
}
