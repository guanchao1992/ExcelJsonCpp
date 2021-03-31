using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using tablegen2.logic;

namespace tablegen2.common
{
    class CreateTableCpp
    {
        static public string CppString1 =
@"#pragma once
#include <string>
#include <vector>
#include <map>
#include <nlohmann/json.hpp>
#include <iostream> 
#include <fstream> 
";

        /*
struct ###TableName###Base
{
	int id;
	std::string name;
	int type;
	int max;
	int sort;
        };
*/

        static public string CppString2 =
        @"
typedef shared_ptr<###TableName###Base> SP###TableName###Base;

struct ###TableName###Datas
{
	###TableName###Datas(const char*path)
	{
		std::ifstream iofile(path);
		nlohmann::json j;
        iofile >> j;
		for (auto it = j.begin(); it != j.end(); it++)
		{
			SP###TableName###Base spTB = make_shared<###TableName###Base>();
";
        /*
        spTB->id = it->at(""id"");
        spTB->name = it->at(""name"");
        spTB->type = it->at(""type"");
        spTB->max = it->at(""max"");
        spTB->sort = it->at(""sort"");
        */

        static public string CppString3 = @"
            _list_datas.push_back(spTB);
			_map_datas[spTB->_id] = spTB;
		}
    }
    static ###TableName###Datas* getInstance()
    {
        static ###TableName###Datas* s_instance = new ###TableName###Datas(""tables/###TableName###.json"");
        return s_instance;
    }
    static std::vector<SP###TableName###Base> getDatas()
    {
        return ###TableName###Datas::getInstance()->_list_datas;
    }
    static SP###TableName###Base getData(int id)
    {
        return ###TableName###Datas::getInstance()->_map_datas[id];
    }
    std::vector<SP###TableName###Base> _list_datas;
    std::map<int, SP###TableName###Base> _map_datas;
};
";
        static public string toFileData(string tableName, List<TableExcelHeader> headers)
        {
            StringBuilder outData = new StringBuilder();
            outData.Append(CppString1);
            StringBuilder descs = new StringBuilder();
            StringBuilder structstr = new StringBuilder();
            structstr.Append("struct ###TableName###Base{\r\n");
            descs.Append(String.Format("/*\r\n"));
            for (int i = 0; i < headers.Count; i++)
            {
                var fieldName = headers[i].FieldName;
                var fieldType = headers[i].FieldType;
                var fieldDesc = headers[i].FieldDesc;
                descs.Append(String.Format("{0} {1}\r\n", fieldName, fieldDesc));
                if (fieldType == "int" ||
                    fieldType == "double" ||
                    fieldType == "float" ||
                    fieldType == "bool"
                    )
                {
                    structstr.Append(String.Format("{0} _{1};\r\n", fieldType, fieldName));
                }
                else if (fieldType == "string")
                {
                    structstr.Append(String.Format("std::string _{0};\r\n", fieldName));
                }
                else if (fieldType == "json")
                {
                    structstr.Append(String.Format("nlohmann::json _{0};\r\n", fieldName));
                }
                else
                {
                    //错误
                    throw new Exception(string.Format(
                        "'{0}'表中数据类型异常，第{1}列\"{2}\"未知的数据类型", tableName, i + 1, fieldType));
                }
            }
            structstr.Append("};");
            descs.Append(String.Format("*/\r\n"));

            outData.Append(descs);
            outData.Append(structstr);
            outData.Append(CppString2);

            for (int i = 0; i < headers.Count; i++)
            {
                var fieldName = headers[i].FieldName;
                var fieldType = headers[i].FieldType;
                var fieldDesc = headers[i].FieldDesc;

                if (fieldType == "json")
                {
                    outData.Append(String.Format("            spTB->_{0} = nlohmann::json::parse(std::string(it->at(\"{0}\")));\r\n", fieldName));
                }
                else
                {
                    outData.Append(String.Format("            spTB->_{0} = it->at(\"{0}\");\r\n", fieldName));
                }

            }
            outData.Append(CppString3);
            outData.Replace("###TableName###", tableName);

            return outData.ToString();
        }
    }
}
