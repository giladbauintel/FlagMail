using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;

namespace FlagEmail
{
    public sealed class MembersDB
    {
        private static volatile MembersDB instance;
        private static object syncRoot = new Object();
        //private SortedDictionary<string, Member> members;
        private DataTable membersTable;

        private MembersDB() 
        {
            initialized = false;

            membersTable = new DataTable("Members");
            DataColumn column;
            //Create Initials column
            column = new DataColumn();
            column.DataType = System.Type.GetType("System.String");
            column.ColumnName = "Initials";
            column.ReadOnly = true;
            column.Unique = true;
            membersTable.Columns.Add(column);

            //Create Full Name column
            column = new DataColumn();
            column.DataType = System.Type.GetType("System.String");
            column.ColumnName = "FullName";
            column.ReadOnly = true;
            column.Unique = true;
            membersTable.Columns.Add(column);

            //Create Username column
            column = new DataColumn();
            column.DataType = System.Type.GetType("System.String");
            column.ColumnName = "Username";
            column.ReadOnly = true;
            column.Unique = true;
            membersTable.Columns.Add(column);

            DataColumn[] PrimaryKeyColumns = new DataColumn[1];
            PrimaryKeyColumns[0] = membersTable.Columns["Initials"];
            membersTable.PrimaryKey = PrimaryKeyColumns;
        }

        public static MembersDB Instance
        {
            get
            {
                if (instance == null)
                {
                    lock (syncRoot)
                    {
                        if (instance == null)
                        {
                            instance = new MembersDB();
                        }
                    }
                }
                return instance;
            }
        }


        public void Add(string initials, string fullName, string username)
        {
            DataRow row = membersTable.NewRow();
            row["Initials"] = initials;
            row["FullName"] = fullName;
            row["Username"] = username;
            membersTable.Rows.Add(row);
        }

        public DataTable getTable()
        {
            return membersTable;
        }

        private bool initialized;
        public bool Initialized
        {
            get {
                return initialized;
            }
            set {
                initialized = value;
            }
        }

        public bool loadJSON(string JSONPath)
        {
            string jsonString;

            //Uri uriResult;
            //bool result = Uri.TryCreate(JSONPath, UriKind.Absolute, out uriResult)
            //    && (uriResult.Scheme == Uri.UriSchemeHttp || uriResult.Scheme == Uri.UriSchemeHttps);
            //
            //if (result)            
            //{
            //    var oauth = new OAuth.Manager();
            //    // the URL to obtain a temporary "request token"
            //    var rtUrl = "https://trello.com/1/OAuthGetRequestToken";
            //    oauth["consumer_key"] = MY_APP_SPECIFIC_KEY;
            //    oauth["consumer_secret"] = MY_APP_SPECIFIC_SECRET;    
            //    oauth.AcquireRequestToken(rtUrl, "POST");
            //    
            //    
            //    WebClient client = new WebClient();
            //    Stream stream = client.OpenRead(JSONPath);
            //    StreamReader reader = new StreamReader(stream);
            //    jsonString = reader.ReadToEnd();           
            //}            
            //else 
            if (File.Exists(JSONPath))
            {                        
                using (StreamReader r = new StreamReader(JSONPath))
                {
                    jsonString = r.ReadToEnd();
                }
            }
            else
            {
                return false;
            }
            
            dynamic json = System.Web.Helpers.Json.Decode(jsonString);
            dynamic members = json.members;

            if (members != null)
            {
                for (int i = 0; i < members.Length; ++i)
                    Add(members[i].initials, members[i].fullName, members[i].username);
            }

            initialized = true;

            return true;
        }
    }
}
