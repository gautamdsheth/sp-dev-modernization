﻿using Newtonsoft.Json;
using System;
using System.Collections.Generic;

namespace SharePointPnP.Modernization.Framework.Entities
{
    /// <summary>
    /// Entity to describe a web part on a wiki or webpart page called from web services
    /// </summary>
    public class WebServiceWebPartEntity
    {
        public WebServiceWebPartEntity()
        {
            this.Properties = new Dictionary<string, string>(StringComparer.InvariantCultureIgnoreCase);
        }

        public string Type { get; set; }


        public Guid Id { get; set; }
        
        public Dictionary<string, string> Properties { get; set; }

        /// <summary>
        /// Shortened web part type name
        /// </summary>
        /// <returns></returns>
        public string TypeShort()
        {
            string name = Type;
            var typeSplit = Type.Split(new string[] { "," }, StringSplitOptions.RemoveEmptyEntries);
            if (typeSplit.Length > 0)
            {
                name = typeSplit[0];
            }

            return $"{name}";
        }

    }
}
