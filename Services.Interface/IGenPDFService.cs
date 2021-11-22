using iTextGenPDF.Api.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace iTextGenPDF.Api.Services.Interface
{
    public interface IGenPDFService
    {
        FileResponseDataBinding CreateCreditDebitNote(Refund creditNoteModel); 
        FileResponseDataBinding CreateTaxInvoice(Payment payment);
        FileResponseDataBinding CreateTrueMoveCDN(Refund refund);

        ResponseModel EncrptPassword(string plainText);
        ResponseModel DecrptPassword(string criperText);
    }
}
