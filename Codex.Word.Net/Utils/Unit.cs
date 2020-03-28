using System;
using System.Globalization;

namespace Codex.Word.Net.Utils
{
    /// <summary>
    /// The basic unit type of page distance.use for <see cref="Unit"/>.
    /// </summary>
    public enum SUnit
    {
        /// <summary>Word Base Unit type, store its 1 times.</summary>
        WordBase = 1,
        /// <summary>Millimeter Unit type, store its 5 times.</summary>
        Milli = 5,
        /// <summary>Centimeter Unit type, store its 10 times.</summary>
        Centi = 10,
        /// <summary>Character Unit type, store its 15 times.</summary>
        Chara = 15,
        /// <summary>Pounds Unit type, store its 20 times.</summary>
        Pound = 20,
        /// <summary>Inches Unit type, store its 30 times.</summary>
        Inch = 30
    }
    
    /// <summary>
    /// Page distance. use for positioning pictures and textboxes in document
    /// </summary>
    /// <example>
    /// <code>
    /// Unit unit = new Unit(100, SUnit.Milli)
    /// </code>
    /// </example>
    /// <remarks>
    /// Longer comments can be associated with a type or member through
    /// the remarks tag.
    /// </remarks>
    public class Unit
    {
        private readonly SUnit _unitType;
        private int _numerical;

        /// <summary>
        /// Initial with float and UnitType string to Word Base Unit
        /// </summary>
        /// <param name="numerical">Numerical of Unit, like 1 in 1cm</param>
        /// <param name="unitType">UnitType of Unit, like cm in 1cm</param>
        public Unit(float numerical, string unitType)
        {
            _unitType = (SUnit)Enum.Parse(typeof(SUnit), unitType);
            Numerical = numerical;
        }
        
        /// <summary>
        /// Initial with float and UnitType to Word Base Unit
        /// </summary>
        /// <param name="numerical">Numerical of Unit, like 1 in 1cm</param>
        /// <param name="unitType">UnitType of Unit, like cm in 1cm</param>
        public Unit(float numerical, SUnit unitType = SUnit.Centi)
        {
            _unitType = unitType;
            Numerical = numerical;
        }

        /// <summary>
        /// Numerical property.
        /// </summary>
        /// <value>
        /// A value tag is used to describe the property value.
        /// </value>
        public float Numerical
        {
            get => _numerical;
            set => _numerical = (int)(value * (int)this._unitType);
        }

        #region Static Method

        /// <summary>
        /// Change the unit of position by adding.
        /// </summary>
        /// <param name="left">The base quantity, the <see cref="SUnit"/> of the result are the same as this parameter</param>
        /// <param name="bias">The incremental</param>
        /// <returns>New Unit with result</returns>
        public static Unit operator +(Unit left, float bias)
        {
            if (left != null)
            {
                Unit result = left;
                result._numerical += (int) (bias * (int) result._unitType);
                return result;
            }

            return null;
        }

        /// <summary>
        /// Change the unit of position by adding
        /// </summary>
        /// <param name="left">The base quantity, the <see cref="SUnit"/> of the result are the same as this parameter</param>
        /// <param name="right">The incremental</param>
        /// <returns>New Unit with result</returns>
        public static Unit operator +(Unit left, Unit right)
        {
            
            if (left != null && right != null)
            {
                Unit result = left;
                result._numerical += right._numerical;
                return result;
            }

            return null;
        }
        
        public static Unit Add(Unit left, Unit right)
        {
            return left+right;
        }

        #endregion

        #region Override
        
        /// <summary>
        /// Provide a way to print the Numerical of Unit
        /// </summary>
        /// <returns>the Numerical of Unit</returns>
        public override string ToString()
        {
            return Numerical.ToString(CultureInfo.CurrentCulture);
        }

        #endregion
    }
}