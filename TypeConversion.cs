using System;
using System.Collections.Generic;
using System.Data;
using System.ComponentModel;
namespace GlobalNetApps.Support.Common
{
    public class TypeConversion
    {
        private static readonly log4net.ILog Log = log4net.LogManager.GetLogger(typeof(TypeConversion));
        /// <summary>
        /// Creates data table for list of model
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="data"></param>
        /// <param name="fileName"></param>
        public DataTable getDataTableforModel<T>(List<T> data)
        {
            DataTable table = new DataTable();
            try
            {
                PropertyDescriptorCollection props = TypeDescriptor.GetProperties(typeof(T));               
                for (int i = 0; i < props.Count; i++)
                {
                    PropertyDescriptor prop = props[i];
                    table.Columns.Add(prop.Name, prop.PropertyType);
                }
                object[] values = new object[props.Count];
                foreach (T item in data)
                {
                    for (int i = 0; i < values.Length; i++)
                    {
                        values[i] = props[i].GetValue(item);
                    }
                    table.Rows.Add(values);
                }
            }
            catch (Exception ex)
            {
                Log.Error(ex.Message);
                throw (ex);
            }
            return table;
        }
    }
}