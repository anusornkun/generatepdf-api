using System;

namespace iTextGenPDF.Api.Models
{
    public class Refund
    {
        public string DocumentName { get; set; }
        public string TypeCode { get; set; }
        public Seller Seller { get; set; }
        public string ActivityDate { get; set; }
        public string DocumentNumber { get; set; }
        public string AccountNumber { get; set; }
        public string SubscriberInfo { get; set; }
        public Buyer Buyer { get; set; }
        public string OriginalReceiptNumber { get; set; }
        public Rnamount RNAmount { get; set; }
        public Corramount CorrAmount { get; set; }
        public Diffamount DiffAmount { get; set; }
        public string AdjustmentReasonDesc { get; set; }
        public string EtaxMedia { get; set; }
        public Etaxinfo[] EtaxInfo { get; set; }
        public string DeliveryMethod { get; set; }
        public string Password { get; set; }
        public string PrintIndicator { get; set; }
        public string ActivityCode { get; set; }
        public string CancelIndicator { get; set; }
        public string GovExtractIndicator { get; set; }
        public string DocumentLanguage { get; set; }
        public string CustomerType { get; set; }
        public string IdentificationType { get; set; }
        public string Identification { get; set; }
        public string BirthDate { get; set; }
        public Mailingaddress MailingAddress { get; set; }
        public string IssueReceiptDate { get; set; }
        public string CalculatedRate { get; set; }
    }

    public class Payment : PaginationModel
    {
        public string DocumentName { get; set; }
        public string TypeCode { get; set; }
        public Seller Seller { get; set; }
        public Buyer Buyer { get; set; }
        public string DepositDate { get; set; }
        public string ReceiptNumber { get; set; }
        public string AccountNumber { get; set; }
        public string PrimResourceVal { get; set; }
        public string PaymentSourceDescription { get; set; }
        public string CreditCardBankAccountNumber { get; set; }
        public Billmonthinfo[] BillMonthInfo { get; set; }
        public string TotalAmount { get; set; }
        public string VatAmount { get; set; }
        public string PaymentAmount { get; set; }
        public Whtrateinfo[] WHTRateInfo { get; set; }
        public string WHTTotalAmount { get; set; }
        public string EtaxMedia { get; set; }
        public Etaxinfo[] EtaxInfo { get; set; }
        public string DeliveryMethod { get; set; }
        public string Password { get; set; }
        public string PrintIndicator { get; set; }
        public string ReasonCode { get; set; }
        public string OriginalReceiptNo { get; set; }
        public string ActivityDate { get; set; }
        public string ActivityCode { get; set; }
        public string OriginalDepositDate { get; set; }
        public string CancelIndicator { get; set; }
        public string GovExtractIndicator { get; set; }
        public string DocumentLanguage { get; set; }
        public string CustomerType { get; set; }
        public string IdentificationType { get; set; }
        public string Identification { get; set; }
        public string BirthDate { get; set; }
        public Mailingaddress MailingAddress { get; set; }
        public string CalculatedRate { get; set; }
        public string TrueMiniKioskNo { get; set; }
        public string TransactionID { get; set; }
        public DetailsPayment[] detailsPayments { get; set; }
    }

    public class Rnamount
    {
        public string Amount { get; set; }
        public string VatAmount { get; set; }
        public string TotalAmount { get; set; }
    }

    public class Corramount
    {
        public string Amount { get; set; }
        public string VatAmount { get; set; }
        public string TotalAmount { get; set; }
    }

    public class Diffamount
    {
        public string Amount { get; set; }
        public string VatAmount { get; set; }
        public string TotalAmount { get; set; }
    }

    public class Seller
    {
        public string SellerName { get; set; }
        public string BranchNo { get; set; }
        public string SellerTaxId { get; set; }
        public string BuildingNumber { get; set; }
        public string BuildingName { get; set; }
        public string Moo { get; set; }
        public string RoomNo { get; set; }
        public string Floor { get; set; }
        public string Soi { get; set; }
        public string SubSoi { get; set; }
        public string StreetName { get; set; }
        public string CitySubDivisionName { get; set; }
        public string CityName { get; set; }
        public string CountrySubDivisionName { get; set; }
        public string PostCode { get; set; }
    }

    public class Buyer
    {
        public string BuyerName { get; set; }
        public string BuyerTaxId { get; set; }
        public string BranchNo { get; set; }
        public string BuildingNumber { get; set; }
        public string BuildingName { get; set; }
        public string Moo { get; set; }
        public string RoomNo { get; set; }
        public string Floor { get; set; }
        public string Soi { get; set; }
        public string SubSoi { get; set; }
        public string StreetName { get; set; }
        public string CitySubDivisionName { get; set; }
        public string CityName { get; set; }
        public string CountrySubDivisionName { get; set; }
        public string PostCode { get; set; }
    }

    public class Billmonthinfo
    {
        public string BillMonth { get; set; }
        public string BeforeVatAmount { get; set; }
        public string DiscountAmount { get; set; }
        public string AfterDiscountAmount { get; set; }
    }

    public class Whtrateinfo
    {
        public string WHTRate { get; set; }
        public string WHTAmount { get; set; }
    }

    public class Etaxinfo
    {
        public string msisdn { get; set; }
        public string type { get; set; }
        public string email { get; set; }
        public string notificationnumber { get; set; }
    }

    public class Mailingaddress
    {
        public string Name { get; set; }
        public string BuildingNumber { get; set; }
        public string BuildingName { get; set; }
        public string Moo { get; set; }
        public string RoomNo { get; set; }
        public string Floor { get; set; }
        public string Soi { get; set; }
        public string SubSoi { get; set; }
        public string StreetName { get; set; }
        public string CitySubDivisionName { get; set; }
        public string CityName { get; set; }
        public string CountrySubDivisionName { get; set; }
        public string PostCode { get; set; }
    }

    public class DetailsPayment
    {
        public string Description { get; set; }
        public string Quantity { get; set; }
        public string UnitPrice { get; set; }
        public string Amount { get; set; }
    }
}
