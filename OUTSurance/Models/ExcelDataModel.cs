using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;


/// <summary>
/// Model to read Values from CSV File.
/// </summary>
/// <history> Created by Mahendher Sangam on 28-July-2017.
/// </history>
namespace OUTSurance.Models
{
    /// <summary>
    /// Model to read Values from  Data.csv File.
    /// </summary>
    public class ExcelDataModel
    {

        /// <summary>
        /// Gets or sets the FirstName.
        /// </summary>
        /// <value>Firstname.</value>
        /// <remarks> The value must be String.
        /// </remarks>
        public string FirstName { get; set; }


        /// <summary>
        /// Gets or sets the LastName.
        /// </summary>
        /// <value>Lastname.</value>
        /// <remarks> The value must be String.
        /// </remarks>
        public string LastName { get; set; }



        /// <summary>
        /// Gets or sets the Address.
        /// </summary>
        /// <value>Lastname.</value>
        /// <remarks> The value must be String.
        /// </remarks>
        public string Address { get; set; }


        
        /// <summary>
        /// Gets or sets the PhoneNumber.
        /// </summary>
        /// <value>PhoneNumber.</value>
        /// <remarks> The value must be String.
        /// </remarks>
        public string PhoneNumber { get; set; } 


    }
}