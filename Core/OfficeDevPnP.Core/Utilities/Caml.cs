using Microsoft.SharePoint.Client;
using System;
using System.Text;

namespace OfficeDevPnP.Core.Utilities {
    /// <summary>
    /// Use this class to build your CAML xml and avoid XML issues.
    /// </summary>
    /// <example>
    /// CAML.ViewQuery(
    ///     CAML.Where(
    ///         CAML.And(
    ///             CAML.Eq(CAML.FieldValue("Project", "Integer", "{0}")),
    ///             CAML.Geq(CAML.FieldValue("StartDate","DateTime", CAML.Today()))
    ///         )
    ///     ),
    ///     CAML.OrderBy(
    ///         new OrderByField("StartDate", false),
    ///         new OrderByField("Title")
    ///     ),
    ///     rowLimit: 5
    /// );
    /// </example>
    public static class CAML {
        const string VIEW_XML_WRAPPER = "<View Scope=\"{0}\"><Query>{1}{2}</Query>{3}<RowLimit>{4}</RowLimit></View>";
        const string FIELD_VALUE = "<FieldRef Name='{0}' {1}/><Value Type='{2}'>{3}</Value>";
        const string FIELD_VALUE_ID = "<FieldRef ID='{0}' {1} /><Value Type='{2}'>{3}</Value>";
        const string WHERE_CLAUSE = "<Where>{0}</Where>";
        const string GENERIC_CLAUSE = "<{0}>{1}</{0}>";
        const string CONDITION_CLAUSE = "<{0}>{1}{2}</{0}>";

        const string VIEW_FIELDS_CLAUSE = "<ViewFields>{0}</ViewFields>";
        const string FIELD_REF_CLAUSE = "<FieldRef Name='{0}'/>";

        public static readonly string Me = "<UserId />";
        public static readonly string Month = "<Month />";
        public static readonly string Now = "<Now />";

        /// <summary>
        /// Creates the &lt;Today /&gt; node.
        /// </summary>
        /// <param name="offset">Time offset from today (+5 days or -5 days, for example).</param>
        /// <returns>Returns &lt;Today /&gt; node based on offset value</returns>
        public static string Today(int? offset = null) {
            if (offset.HasValue)
                return $"<Today Offset='{offset.Value}' />";
            return "<Today />";
        }

        /// <summary>
        /// Root &lt;View&gt; and &lt;Query&gt; nodes.
        /// </summary>
        /// <param name="whereClause">&lt;Where&gt; node.</param>
        /// <param name="orderByClause">&lt;OrderBy&gt; node.</param>
        /// <param name="rowLimit">&lt;RowLimit&gt; node.</param>
        /// <returns>String to be used in CAML queries</returns>
        public static string ViewQuery(string whereClause = "", string orderByClause = "", int rowLimit = 100) {
            return CAML.ViewQuery(ViewScope.DefaultValue, whereClause, orderByClause, string.Empty, rowLimit);
        }

        /// <summary>
        /// Root &lt;View&gt; and &lt;Query&gt; nodes.
        /// </summary>
        /// <param name="scope">View scope</param>
        /// <param name="whereClause">&lt;Where&gt; node.</param>
        /// <param name="viewFields">&lt;ViewFields&gt; node.</param>
        /// <param name="orderByClause">&lt;OrderBy&gt; node.</param>
        /// <param name="rowLimit">&lt;RowLimit&gt; node.</param>
        /// <returns>String to be used in CAML queries</returns>
        public static string ViewQuery(ViewScope scope, string whereClause = "", string orderByClause = "", string viewFields = "", int rowLimit = 100) {
            string viewScopeStr = scope == ViewScope.DefaultValue ? string.Empty : scope.ToString();
            return string.Format(VIEW_XML_WRAPPER, viewScopeStr, whereClause, orderByClause, viewFields, rowLimit);
        }

        /// <summary>
        /// Creates both a &lt;FieldRef&gt; and &lt;Value&gt; nodes combination for Where clauses.
        /// </summary>
        /// <param name="fieldName">Name of the field</param>
        /// <param name="fieldValueType">Value type of the field</param>
        /// <param name="value">Value of the field</param>
        /// <param name="additionalFieldRefParams">Additional FieldRef Parameters</param>
        /// <returns>Returns FieldValue string to be used in CAML queries</returns>
        public static string FieldValue(string fieldName, string fieldValueType, string value, string additionalFieldRefParams = "") {
            return string.Format(FIELD_VALUE, fieldName, additionalFieldRefParams, fieldValueType, value);
        }

        /// <summary>
        /// Creates both a &lt;FieldRef&gt; and &lt;Value&gt; nodes combination for Where clauses.
        /// </summary>
        /// <param name="fieldId">Id of the field</param>
        /// <param name="fieldValueType">Value type of the field</param>
        /// <param name="value">Value of the field</param>
        /// <param name="additionalFieldRefParams">Additional FieldRef Parameters</param>
        /// <returns>Returns FieldValue string to be used in CAML queries</returns>
        public static string FieldValue(Guid fieldId, string fieldValueType, string value, string additionalFieldRefParams = "") {
            return string.Format(FIELD_VALUE_ID, fieldId.ToString(), additionalFieldRefParams, fieldValueType, value);
        }

        /// <summary>
        /// Creates a &lt;FieldRef&gt; node for ViewFields clause
        /// </summary>
        /// <param name="fieldName">Name of the field</param>
        /// <returns>Returns FieldRef string to be used in CAML queries</returns>
        public static string FieldRef(string fieldName) {
            return string.Format(FIELD_REF_CLAUSE, fieldName);
        }

        /// <summary>
        /// Creates &lt;OrederBy&gt; node for sorting by field
        /// </summary>
        /// <param name="fieldRefs">Field References</param>
        /// <returns>Returns string to be used in CAML queries</returns>
        public static string OrderBy(params OrderByField[] fieldRefs) {
            var sb = new StringBuilder();
            foreach (var field in fieldRefs){
                sb.Append(field.ToString());
            }
            return string.Format(GENERIC_CLAUSE, CamlClauses.OrderBy, sb.ToString());
        }

        /// <summary>
        /// Creates &lt;Where&gt; node for Where clause
        /// </summary>
        /// <param name="conditionClause">The Clause condition</param>
        /// <returns>Returns string to be used in CAML queries</returns>
        public static string Where(string conditionClause) {
            return string.Format(GENERIC_CLAUSE, CamlClauses.Where, conditionClause);
        }

        /// <summary>
        /// Creates &lt;ViewFields&gt; node for ViewFields clause
        /// </summary>
        /// <param name="fieldRefs">Field References</param>
        /// <returns>Returns string to be used in CAML queries</returns>
        public static string ViewFields(params string[] fieldRefs) {
            string refs = string.Empty;

            foreach (var refField in fieldRefs) {
                refs += refField;
            }
            return string.Format(VIEW_FIELDS_CLAUSE, refs);
        }

        #region Conditions
        /// <summary>
        /// Creates &lt;And&gt; node 
        /// </summary>
        /// <param name="clause1">Clause</param>
        /// <param name="conditionClauses">Clause Condition</param>
        /// <returns>Returns Condition string to be used in CAML queries</returns>
        public static string And(string clause1, params string[] conditionClauses) {
            return Condition(CamlConditions.And, clause1, conditionClauses);
        }

        /// <summary>
        /// Creates &lt;Or&gt; node
        /// </summary>
        /// <param name="clause1">Clause</param>
        /// <param name="conditionClauses">Clause Condition</param>
        /// <returns>Returns Condition string to be used in CAML queries</returns>
        public static string Or(string clause1, params string[] conditionClauses) {
            return Condition(CamlConditions.Or, clause1, conditionClauses);
        } 
        #endregion

        #region Comparisons
        /// <summary>
        /// Creates &lt;BeginsWith&gt; node for Comparison
        /// </summary>
        /// <param name="fieldValue">Value of the field</param>
        /// <returns>Returns Comparison string to be used in CAML queries</returns>
        public static string BeginsWith(string fieldValue) {
            return Comparison(CamlComparisions.BeginsWith, fieldValue);
        }
        /// <summary>
        /// Creates &lt;Contains&gt; node for Comparison
        /// </summary>
        /// <param name="fieldValue">Value of the field</param>
        /// <returns>Returns Comparison string to be used in CAML queries</returns>
        public static string Contains(string fieldValue) {
            return Comparison(CamlComparisions.Contains, fieldValue);
        }
        /// <summary>
        /// Creates &lt;DateRangesOverlap&gt; node for Comparison
        /// </summary>
        /// <param name="fieldValue">Value of the field</param>
        /// <returns>Returns Comparison string to be used in CAML queries</returns>
        public static string DateRangesOverlap(string fieldValue) {
            return Comparison(CamlComparisions.DateRangesOverlap, fieldValue);
        }
        /// <summary>
        /// Creates &lt;Eq&gt; node for Comparison
        /// </summary>
        /// <param name="fieldValue">Value of the field</param>
        /// <returns>Returns Comparison string to be used in CAML queries</returns>
        public static string Eq(string fieldValue) {
            return Comparison(CamlComparisions.Eq, fieldValue);
        }
        /// <summary>
        /// Creates &lt;Geq&gt; node for Comparison
        /// </summary>
        /// <param name="fieldValue">Value of the field</param>
        /// <returns>Returns Comparison string to be used in CAML queries</returns>
        public static string Geq(string fieldValue) {
            return Comparison(CamlComparisions.Geq, fieldValue);
        }
        /// <summary>
        /// Creates &lt;Gt&gt; node for Comparison
        /// </summary>
        /// <param name="fieldValue">Value of the field</param>
        /// <returns>Returns Comparison string to be used in CAML queries</returns>
        public static string Gt(string fieldValue) {
            return Comparison(CamlComparisions.Gt, fieldValue);
        }
        /// <summary>
        /// Creates &lt;In&gt; node for Comparison
        /// </summary>
        /// <param name="fieldValue">Value of the field</param>
        /// <returns>Returns Comparison string to be used in CAML queries</returns>
        public static string In(string fieldValue) {
            return Comparison(CamlComparisions.In, fieldValue);
        }
        /// <summary>
        /// Creates &lt;Includes&gt; node for Comparison
        /// </summary>
        /// <param name="fieldValue">Value of the field</param>
        /// <returns>Returns Comparison string to be used in CAML queries</returns>
        public static string Includes(string fieldValue) {
            return Comparison(CamlComparisions.Includes, fieldValue);
        }
        /// <summary>
        /// Creates &lt;IsNotNull&gt; node for Comparison
        /// </summary>
        /// <param name="fieldValue">Value of the field</param>
        /// <returns>Returns Comparison string to be used in CAML queries</returns>
        public static string IsNotNull(string fieldValue) {
            return Comparison(CamlComparisions.IsNotNull, fieldValue);
        }
        /// <summary>
        /// Creates &lt;IsNull&gt; node for Comparison
        /// </summary>
        /// <param name="fieldValue">Value of the field</param>
        /// <returns>Returns Comparison string to be used in CAML queries</returns>
        public static string IsNull(string fieldValue) {
            return Comparison(CamlComparisions.IsNull, fieldValue);
        }
        /// <summary>
        /// Creates &lt;Leq&gt; node for Comparison
        /// </summary>
        /// <param name="fieldValue">Value of the field</param>
        /// <returns>Returns Comparison string to be used in CAML queries</returns>
        public static string Leq(string fieldValue) {
            return Comparison(CamlComparisions.Leq, fieldValue);
        }
        /// <summary>
        /// Creates &lt;Lt&gt; node for Comparison
        /// </summary>
        /// <param name="fieldValue">Value of the field</param>
        /// <returns>Returns Comparison string to be used in CAML queries</returns>
        public static string Lt(string fieldValue) {
            return Comparison(CamlComparisions.Lt, fieldValue);
        }
        /// <summary>
        /// Creates &lt;Neq&gt; node for Comparison
        /// </summary>
        /// <param name="fieldValue">Value of the field</param>
        /// <returns>Returns Comparison string to be used in CAML queries</returns>
        public static string Neq(string fieldValue) {
            return Comparison(CamlComparisions.Neq, fieldValue);
        }
        /// <summary>
        /// Creates &lt;NotIncludes&gt; node for Comparison
        /// </summary>
        /// <param name="fieldValue">Value of the field</param>
        /// <returns>Returns Comparison string to be used in CAML queries</returns>
        public static string NotIncludes(string fieldValue) {
            return Comparison(CamlComparisions.NotIncludes, fieldValue);
        }
        #endregion

        static string Comparison(CamlComparisions comparison, string fieldValue) {
            return string.Format(GENERIC_CLAUSE, comparison, fieldValue);
        }

        static string Condition(CamlConditions condition, string clause1, params string[] comparisonClauses) {
            var formattedString = clause1;

            foreach (var clause in comparisonClauses) {
                formattedString = string.Format(CONDITION_CLAUSE, condition, formattedString, clause);
            }

            return formattedString;
        }

        enum CamlComparisions {
            BeginsWith, Contains, DateRangesOverlap,
            Eq, Geq, Gt, In, Includes, IsNotNull, IsNull,
            Leq, Lt, Neq, NotIncludes
        }
        enum CamlConditions { And, Or }
        enum CamlClauses { Where, OrderBy, GroupBy }

    }

    /// <summary>
    /// Class is used to order the data by field.
    /// </summary>
    public class OrderByField {
        const string ORDERBY_FIELD = "<FieldRef Name='{0}' Ascending='{1}' />";
        /// <summary>
        /// Constructor for OrderByField class
        /// </summary>
        /// <param name="name">Name of the field</param>
        public OrderByField(string name) : this(name, true) { }
        /// <summary>
        /// Constructor for OrderByField class
        /// </summary>
        /// <param name="name">Name of the field</param>
        /// <param name="ascending">If we want to order in ascending order</param>
        public OrderByField(string name, bool ascending) {
            Name = name;
            Ascending = ascending;
        }
        /// <summary>
        /// Gets or sets the name of the field
        /// </summary>
        public string Name { get; set; }
        /// <summary>
        /// Gets or sets the Ascending order flag
        /// </summary>
        public bool Ascending { get; set; }
        /// <summary>
        /// OrderByField string
        /// </summary>
        /// <returns>Returns string</returns>
        public override string ToString() {
            return string.Format(ORDERBY_FIELD, Name, Ascending.ToString().ToUpper());
        }
    }
}
