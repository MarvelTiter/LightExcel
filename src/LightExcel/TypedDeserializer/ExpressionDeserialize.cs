using LightExcel;
using LightExcel.Attributes;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace LightExcel.TypedDeserializer
{
    /// <summary>
    /// 从LightOrm中移植修改
    /// </summary>
    internal static class ExpressionDeserialize<T>
    {
        private static readonly MethodInfo DataRecord_GetInt16 = typeof(IExcelDataReader).GetMethod("GetInt16", [typeof(int)])!;
        private static readonly MethodInfo DataRecord_GetInt32 = typeof(IExcelDataReader).GetMethod("GetInt32", [typeof(int)])!;
        private static readonly MethodInfo DataRecord_GetInt64 = typeof(IExcelDataReader).GetMethod("GetInt64", [typeof(int)])!;
        private static readonly MethodInfo DataRecord_GetDouble = typeof(IExcelDataReader).GetMethod("GetDouble", [typeof(int)])!;
        private static readonly MethodInfo DataRecord_GetDecimal = typeof(IExcelDataReader).GetMethod("GetDecimal", [typeof(int)])!;
        private static readonly MethodInfo DataRecord_GetBoolean = typeof(IExcelDataReader).GetMethod("GetBoolean", [typeof(int)])!;
        private static readonly MethodInfo DataRecord_GetDateTime = typeof(IExcelDataReader).GetMethod("GetDateTime", [typeof(int)])!;
        private static readonly MethodInfo DataRecord_GetValue = typeof(IExcelDataReader).GetMethod("GetValue", [typeof(int)])!;
        private static readonly MethodInfo IsNullOrEmpty = typeof(IExcelDataReader).GetMethod("IsNullOrEmpty", [typeof(int)])!;
        readonly static Dictionary<Type, MethodInfo> typeMapMethod = new Dictionary<Type, MethodInfo>(37)
        {
            [typeof(short)] = DataRecord_GetInt16,
            [typeof(ushort)] = DataRecord_GetInt16,
            [typeof(int)] = DataRecord_GetInt32,
            [typeof(uint)] = DataRecord_GetInt32,
            [typeof(long)] = DataRecord_GetInt64,
            [typeof(ulong)] = DataRecord_GetInt64,
            [typeof(double)] = DataRecord_GetDouble,
            [typeof(decimal)] = DataRecord_GetDecimal,
            [typeof(bool)] = DataRecord_GetBoolean,
            [typeof(DateTime)] = DataRecord_GetDateTime,
        };

        static Func<IExcelDataReader, object>? handler;


        public static T Deserialize(IExcelDataReader reader)
        {
            handler ??= BuildFunc<T>(reader, CultureInfo.CurrentCulture, false);
            return (T)handler.Invoke(reader);
        }

        /// <summary>
        /// record => {
        ///     return new T {
        ///         Member0 = (memberType)record.Get_XXX[0],
        ///         Member1 = (memberType)record.Get_XXX[1],
        ///         Member2 = (memberType)record.Get_XXX[2],
        ///         Member3 = (memberType)record.Get_XXX[3],
        ///         Member4 = record.IsDBNull(4) ? default(memberType) : (memberType)record.Get_XXX[4],
        ///     }
        /// }
        /// </summary>
        /// <typeparam name="Target"></typeparam>
        /// <param name="RecordInstance"></param>
        /// <param name="Culture"></param>
        /// <param name="MustMapAllProperties"></param>
        /// <returns></returns>
        private static Func<IExcelDataReader, object> BuildFunc<Target>(IExcelDataReader RecordInstance, CultureInfo Culture, bool MustMapAllProperties)
        {
            ParameterExpression recordInstanceExp = Expression.Parameter(typeof(IExcelDataReader), "Record");
            Type TargetType = typeof(Target);
            Expression? Body = default;

            // 元组处理
            if (TargetType.FullName?.StartsWith("System.Tuple`") ?? false)
            {
                ConstructorInfo[] Constructors = TargetType.GetConstructors();
                if (Constructors.Count() != 1)
                    throw new ArgumentException("Tuple must have one Constructor");
                var Constructor = Constructors[0];

                var Parameters = Constructor.GetParameters();
                if (Parameters.Length > 7)
                    throw new NotSupportedException("Nested Tuples are not supported");

                Expression[] TargetValueExpressions = new Expression[Parameters.Length];
                for (int Ordinal = 0; Ordinal < Parameters.Length; Ordinal++)
                {
                    var ParameterType = Parameters[Ordinal].ParameterType;
                    if (Ordinal >= RecordInstance.FieldCount)
                    {
                        if (MustMapAllProperties) { throw new ArgumentException("Tuple has more fields than the DataReader"); }
                        TargetValueExpressions[Ordinal] = Expression.Default(ParameterType);
                    }
                    else
                    {
                        TargetValueExpressions[Ordinal] = GetTargetValueExpression(
                                                        RecordInstance,
                                                        Culture,
                                                        recordInstanceExp,
                                                        Ordinal,
                                                        ParameterType);
                    }
                }
                Body = Expression.New(Constructor, TargetValueExpressions);
            }
            // 基础类型处理 eg: IEnumable<int>  IEnumable<string>
            else if (TargetType.IsElementaryType())
            {
                const int Ordinal = 0;
                Expression TargetValueExpression = GetTargetValueExpression(
                                                        RecordInstance,
                                                        Culture,
                                                        recordInstanceExp,
                                                        Ordinal,
                                                        TargetType);

                UnaryExpression converted = Expression.Convert(TargetValueExpression, typeof(object));
                Body = Expression.Block(converted);
            }
            // 其他
            else
            {
                SortedDictionary<int, MemberBinding> Bindings = new SortedDictionary<int, MemberBinding>();
                // 字段处理 Field
                foreach (FieldInfo TargetMember in TargetType.GetFields(BindingFlags.Public | BindingFlags.Instance))
                {
                    Action work = delegate
                    {
                        for (int Ordinal = 0; Ordinal < RecordInstance.FieldCount; Ordinal++)
                        {
                            //Check if the RecordFieldName matches the TargetMember
                            if (MemberMatchesName(TargetMember, RecordInstance.GetName(Ordinal)))
                            {
                                Expression TargetValueExpression = GetTargetValueExpression(
                                                                        RecordInstance,
                                                                        Culture,
                                                                        recordInstanceExp,
                                                                        Ordinal,
                                                                        TargetMember.FieldType);

                                //Create a binding to the target member
                                MemberAssignment BindExpression = Expression.Bind(TargetMember, TargetValueExpression);
                                Bindings.Add(Ordinal, BindExpression);
                                return;
                            }
                        }
                        //If we reach this code the targetmember did not get mapped
                        if (MustMapAllProperties)
                        {
                            throw new ArgumentException(string.Format("TargetField {0} is not matched by any field in the DataReader", TargetMember.Name));
                        }
                    };
                    work();
                }
                // 属性处理 Property
                foreach (PropertyInfo TargetMember in TargetType.GetProperties(BindingFlags.Public | BindingFlags.Instance))
                {
                    if (TargetMember.CanWrite)
                    {
                        Action work = delegate
                        {
                            for (int Ordinal = 0; Ordinal < RecordInstance.FieldCount; Ordinal++)
                            {
                                //Check if the RecordFieldName matches the TargetMember
                                if (MemberMatchesName(TargetMember, RecordInstance.GetName(Ordinal)))
                                {
                                    Expression TargetValueExpression = GetTargetValueExpression(
                                                                            RecordInstance,
                                                                            Culture,
                                                                            recordInstanceExp,
                                                                            Ordinal,
                                                                            TargetMember.PropertyType);

                                    //Create a binding to the target member
                                    MemberAssignment BindExpression = Expression.Bind(TargetMember, TargetValueExpression);
                                    Bindings.Add(Ordinal, BindExpression);
                                    return;
                                }
                            }
                            //If we reach this code the targetmember did not get mapped
                            if (MustMapAllProperties)
                            {
                                throw new ArgumentException(string.Format("TargetProperty {0} is not matched by any Field in the DataReader", TargetMember.Name));
                            }
                        };
                        work();
                    }
                }

                Body = Expression.MemberInit(Expression.New(TargetType), Bindings.Values);

            }
            //Compile as Delegate
            var lambdaExp = Expression.Lambda<Func<IExcelDataReader, object>>(Body, recordInstanceExp);
            return lambdaExp.Compile();
        }

        private static bool MemberMatchesName(MemberInfo Member, string Name)
        {
            string FieldnameAttribute = GetColumnNameAttribute();
            return FieldnameAttribute.ToLower() == Name.ToLower() || Member.Name.ToLower() == Name.ToLower();

            string GetColumnNameAttribute()
            {
                if (Member.GetCustomAttributes(typeof(ExcelColumnAttribute), true).Count() > 0)
                {
                    return ((ExcelColumnAttribute)Member.GetCustomAttributes(typeof(ExcelColumnAttribute), true)[0]).Name ?? string.Empty;
                }
                else if (Member.IsDefined(typeof(ExcelColumnAttribute), true))
                {
                    return (Member.GetCustomAttribute(typeof(ExcelColumnAttribute), true) as ExcelColumnAttribute)?.Name ?? string.Empty;
                }
                else
                {
                    return string.Empty;
                }
            }
        }

        private static Expression GetTargetValueExpression(
            IExcelDataReader RecordInstance,
            CultureInfo Culture,
            ParameterExpression recordInstanceExp,
            int Ordinal,
            Type TargetMemberType)
        {
            var needConvert = GetRecordFieldExpression(recordInstanceExp, Ordinal, TargetMemberType, out var RecordFieldExpression);
            Expression ConvertedRecordFieldExpression = needConvert ? GetConversionExpression(typeof(string), RecordFieldExpression, TargetMemberType, Culture)
                : RecordFieldExpression;
            MethodCallExpression NullCheckExpression = GetNullCheckExpression(recordInstanceExp, Ordinal);

            Expression TargetValueExpression = Expression.Condition(
            NullCheckExpression,
            Expression.Default(TargetMemberType),
            ConvertedRecordFieldExpression,
            TargetMemberType
            );
            return TargetValueExpression;
        }

        private static bool GetRecordFieldExpression(ParameterExpression recordInstanceExp, int Ordinal, Type RecordFieldType, out Expression expression)
        {
            //MethodInfo GetValueMethod = default(MethodInfo);
            var has = typeMapMethod.TryGetValue(RecordFieldType, out var GetValueMethod);
            if (GetValueMethod == null)
                GetValueMethod = DataRecord_GetValue;

            expression = Expression.Call(recordInstanceExp, GetValueMethod, Expression.Constant(Ordinal, typeof(int)));

            return !has;
        }

        private static MethodCallExpression GetNullCheckExpression(ParameterExpression RecordInstance, int Ordinal)
        {
            MethodCallExpression NullCheckExpression = Expression.Call(RecordInstance, IsNullOrEmpty, Expression.Constant(Ordinal, typeof(int)));
            return NullCheckExpression;
        }

        private static Expression GetConversionExpression(Type SourceType, Expression SourceExpression, Type TargetType, CultureInfo Culture)
        {
            Expression TargetExpression;
            if (ReferenceEquals(TargetType, SourceType))
            {
                TargetExpression = SourceExpression;
            }
            else if (ReferenceEquals(SourceType, typeof(string)))
            {
                TargetExpression = GetParseExpression(SourceExpression, TargetType, Culture);
            }
            else if (ReferenceEquals(TargetType, typeof(string)))
            {
                TargetExpression = Expression.Call(SourceExpression, SourceType.GetMethod("ToString", Type.EmptyTypes)!);
            }
            else if (ReferenceEquals(TargetType, typeof(bool)))
            {
                MethodInfo ToBooleanMethod = typeof(Convert).GetMethod("ToBoolean", [SourceType])!;
                TargetExpression = Expression.Call(ToBooleanMethod, SourceExpression);
            }
            else if (ReferenceEquals(SourceType, typeof(byte[])))
            {
                throw new NotSupportedException();
            }
            else
            {
                TargetExpression = Expression.Convert(SourceExpression, TargetType);
            }
            return TargetExpression;
        }


        private static Expression GetParseExpression(Expression SourceExpression, Type TargetType, CultureInfo Culture)
        {
            Type UnderlyingType = GetUnderlyingType(TargetType);
            if (UnderlyingType.IsEnum)
            {
                MethodCallExpression ParsedEnumExpression = GetEnumParseExpression(SourceExpression, UnderlyingType);
                //Enum.Parse returns an object that needs to be unboxed
                return Expression.Unbox(ParsedEnumExpression, TargetType);
            }
            else
            {
                Expression? ParseExpression = default;
                switch (UnderlyingType.FullName)
                {
                    case "System.Byte":
                    case "System.UInt16":
                    case "System.UInt32":
                    case "System.UInt64":
                    case "System.SByte":
                    case "System.Int16":
                    case "System.Int32":
                    case "System.Int64":
                    case "System.Double":
                    case "System.Decimal":
                        ParseExpression = GetNumberParseExpression(SourceExpression, UnderlyingType, Culture);
                        break;
                    case "System.DateTime":
                        ParseExpression = GetDateTimeParseExpression(SourceExpression, UnderlyingType, Culture);
                        break;
                    case "System.Boolean":
                    case "System.Char":
                        ParseExpression = GetGenericParseExpression(SourceExpression, UnderlyingType);
                        break;
                    default:
                        throw new ArgumentException(string.Format("Conversion from {0} to {1} is not supported", "String", TargetType));
                }
                if (Nullable.GetUnderlyingType(TargetType) == null)
                {
                    return ParseExpression;
                }
                else
                {
                    //Convert to nullable if necessary
                    return Expression.Convert(ParseExpression, TargetType);
                }
            }
            Expression GetGenericParseExpression(Expression sourceExpression, Type type)
            {
                MethodInfo ParseMetod = type.GetMethod("Parse", [typeof(string)])!;
                MethodCallExpression CallExpression = Expression.Call(ParseMetod, [sourceExpression]);
                return CallExpression;
            }
            Expression GetDateTimeParseExpression(Expression sourceExpression, Type type, CultureInfo culture)
            {
                MethodInfo ParseMetod = type.GetMethod("Parse", [typeof(string), typeof(DateTimeFormatInfo)])!;
                ConstantExpression ProviderExpression = Expression.Constant(culture.DateTimeFormat, typeof(DateTimeFormatInfo));
                MethodCallExpression CallExpression = Expression.Call(ParseMetod, [sourceExpression, ProviderExpression]);
                return CallExpression;
            }

            MethodCallExpression GetEnumParseExpression(Expression sourceExpression, Type type)
            {
                //Get the MethodInfo for parsing an Enum
                MethodInfo EnumParseMethod = typeof(Enum).GetMethod("Parse", [typeof(Type), typeof(string), typeof(bool)])!;
                ConstantExpression TargetMemberTypeExpression = Expression.Constant(type);
                ConstantExpression IgnoreCase = Expression.Constant(true, typeof(bool));
                //Create an expression the calls the Parse method
                MethodCallExpression CallExpression = Expression.Call(EnumParseMethod, [TargetMemberTypeExpression, sourceExpression, IgnoreCase]);
                return CallExpression;
            }

            MethodCallExpression GetNumberParseExpression(Expression sourceExpression, Type type, CultureInfo culture)
            {
                MethodInfo ParseMetod = type.GetMethod("Parse", [typeof(string), typeof(NumberFormatInfo)])!;
                ConstantExpression ProviderExpression = Expression.Constant(culture.NumberFormat, typeof(NumberFormatInfo));
                MethodCallExpression CallExpression = Expression.Call(ParseMetod, [sourceExpression, ProviderExpression]);
                return CallExpression;
            }
        }

        private static Type GetUnderlyingType(Type targetType)
        {
            return Nullable.GetUnderlyingType(targetType) ?? targetType;
        }
    }
    public static class Ex
    {
        /// <summary>
        /// 检查是否为基础类型
        /// </summary>
        /// <param name="t"></param>
        /// <returns></returns>
        public static bool IsElementaryType(this Type t)
        {
            return ElementaryTypes.Contains(t);
        }
        readonly static HashSet<Type> ElementaryTypes = LoadElementaryTypes();
        private static HashSet<Type> LoadElementaryTypes()
        {
            HashSet<Type> TypeSet = new HashSet<Type>()
            {
                    typeof(string),
                    typeof(byte),
                    typeof(sbyte),
                    typeof(short),
                    typeof(int),
                    typeof(long),
                    typeof(ushort),
                    typeof(uint),
                    typeof(ulong),
                    typeof(float),
                    typeof(double),
                    typeof(decimal),
                    typeof(DateTime),
                    typeof(Guid),
                    typeof(bool),
                    typeof(TimeSpan),
                    typeof(byte?),
                    typeof(sbyte?),
                    typeof(short?),
                    typeof(int?),
                    typeof(long?),
                    typeof(ushort?),
                    typeof(uint?),
                    typeof(ulong?),
                    typeof(float?),
                    typeof(double?),
                    typeof(decimal?),
                    typeof(DateTime?),
                    typeof(Guid?),
                    typeof(bool?),
                    typeof(TimeSpan?)
                };
            return TypeSet;
        }
    }
}
