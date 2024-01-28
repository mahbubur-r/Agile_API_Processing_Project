public class AgreementBids
{
    public string _id { get; set; }
    public int agreement_title_id { get; set; }
    public string agreement_title { get; set; }
    public string project_information { get; set; }
    public string employee_name { get; set; }
    public string provider_name { get; set; }
    public string contactperson { get; set; }
    public string externalperson { get; set; }
    public string rate { get; set; }
    public string dateuntil { get; set; }
    public string notes { get; set; }
    public string document { get; set; }
    public string status { get; set; }
    public int __v { get; set; }

    // New property added for provider
    public string Provider { get; set; }
}