using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.ComponentModel;
using System.Globalization;
using System.Linq.Expressions;
using System.Diagnostics;

namespace IISHelper.Extensions {
    public static class ObjectExtensions {
        public static T CloneObject<T>(this T sourceObject) {
            Type t = sourceObject.GetType();
            PropertyInfo[] properties = t.GetProperties();
            Object p = t.InvokeMember("", System.Reflection.BindingFlags.CreateInstance, null, sourceObject, null);

            foreach (PropertyInfo pi in properties) {
                if (pi.CanWrite) {
                    pi.SetValue(p, pi.GetValue(sourceObject, null), null);
                }
            }
            return (T)Convert.ChangeType(p, typeof(T));
        }


        public static T GetAttribute<T>(this MemberInfo member, bool isRequired)
    where T : Attribute {
            var attribute = member.GetCustomAttributes(typeof(T), false).SingleOrDefault();

            if (attribute == null && isRequired) {
                throw new ArgumentException(
                    string.Format(
                        CultureInfo.InvariantCulture,
                        "The {0} attribute must be defined on member {1}",
                        typeof(T).Name,
                        member.Name));
            }

            return (T)attribute;
        }

        public static string GetPropertyDisplayName<T>(Expression<Func<T, object>> propertyExpression) {
            var memberInfo = GetPropertyInformation(propertyExpression.Body);
            if (memberInfo == null) {
                throw new ArgumentException(
                    "No property reference expression was found.",
                    "propertyExpression");
            }

            var attr = memberInfo.GetAttribute<DisplayNameAttribute>(false);
            if (attr == null) {
                return memberInfo.Name;
            }

            return attr.DisplayName;
        }

        public static MemberInfo GetPropertyInformation(Expression propertyExpression) {
            Debug.Assert(propertyExpression != null, "propertyExpression != null");
            MemberExpression memberExpr = propertyExpression as MemberExpression;
            if (memberExpr == null) {
                UnaryExpression unaryExpr = propertyExpression as UnaryExpression;
                if (unaryExpr != null && unaryExpr.NodeType == ExpressionType.Convert) {
                    memberExpr = unaryExpr.Operand as MemberExpression;
                }
            }

            if (memberExpr != null && memberExpr.Member.MemberType == MemberTypes.Property) {
                return memberExpr.Member;
            }

            return null;
        }
    }
}
