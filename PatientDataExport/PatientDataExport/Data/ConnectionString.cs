using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace PatientDataExport.Data
{
    public class ConnectionString
    {
        //保留 用于选择连接的数据库
        public System.Data.SqlClient.SqlConnectionStringBuilder ConStr()
        {
            //<add name="medbaseEntities" 
            //connectionString="metadata=res://*/Data.PatientData.csdl|res://*/Data.PatientData.ssdl|res://*/Data.PatientData.msl;
            //provider=System.Data.SqlClient;provider connection string=&quot;
            //data source=192.168.1.161;
            //initial catalog=medbase;
            //user id=sa;password=@Zhangkai851983;
            //MultipleActiveResultSets=True;
            //App=EntityFramework&quot;" 
            //providerName="System.Data.EntityClient" />
            System.Data.SqlClient.SqlConnectionStringBuilder myConnectionString = new System.Data.SqlClient.SqlConnectionStringBuilder();
            myConnectionString.ConnectionString = @"metadata=res://*/Data.PatientData.csdl|res://*/Data.PatientData.ssdl|res://*/Data.PatientData.msl;provider=System.Data.SqlClient;provider connection string=&quot;data source=192.168.1.161;initial catalog=medbase;user id=sa;password=@Zhangkai851983;MultipleActiveResultSets=True;App=EntityFramework&quot;";
            return myConnectionString;
        }
    }
}
