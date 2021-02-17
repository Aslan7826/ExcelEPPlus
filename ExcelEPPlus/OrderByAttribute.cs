using System;
using System.ComponentModel.DataAnnotations;

namespace ExcelEPPlus
{
    sealed public class OrderByAttribute : ValidationAttribute, IComparable<OrderByAttribute>
    {
        private int _OrderBy
        {
            get;
            set;
        }


        public OrderByAttribute(int OrderBy)
        {
            _OrderBy = OrderBy;
        }


        public int CompareTo(OrderByAttribute other)
        {
            return this._OrderBy.CompareTo(other._OrderBy);
            //throw new NotImplementedException();
        }
    }
}