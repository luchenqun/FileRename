#include <iostream>
#include <algorithm>
#include <map>
#include <io.h>
#include <string>
#include <cctype>
#include <Windows.h>   
#include "BasicExcel.h"
#include "DataConst.h"

using namespace std;
using namespace YExcel;

void getFiles( string, vector<string>&);
void gb2312ToUnicode(const string& src, wstring& result);
string getSpell(const string fileName);
void checkFileName(const string fileName);
int Between(int var, int lower, int upper);// �ж�var��ֵ�Ƿ�������֮��
string GetPinyinHead(char HI, char LO);// ����һ�����ֵ�ƴ������ĸ

// ȫ�ֱ��������ڴ�Ŵ�����Ϣ
vector<string>	errorInfo;

int main()
{
	//	freopen("out.txt", "w", stdout);

	vector<string>	files;

	BasicExcel excel;
	excel.New(1);
	excel.RenameWorksheet("Sheet1", "FileRename");
	BasicExcelWorksheet* sheet = excel.GetWorksheet("FileRename");

	getFiles("ԭʼ", files);// ��ȡԭ�ļ������е��ļ�

	int excelIndex = 0;
	int fileNums = files.size();
	for(int fileIndex=0; fileIndex<fileNums; ++fileIndex)
	{
		string sChinese = files[fileIndex];   //   ������ַ���
		string pinYin;

		sChinese.erase(0, 5);	// ɾ�� "ԭʼ\" ��7���ַ�
		transform(sChinese.begin(), sChinese.end(), 
			sChinese.begin(), toupper);// ���������е���ĸ��СдתΪ��д

		checkFileName(sChinese);// ����ļ����Ƿ�Ϸ�
		pinYin = getSpell(sChinese);// ��ȡ�ú��ֵ�ƴ������ĸ

		sChinese.erase(sChinese.length()-4, sChinese.length());	// ɾ�� ".bmp" ��4���ַ�

		//		cout << sChinese <<" " << pinYin << endl; 

		wstring wstrUnicode;  
		gb2312ToUnicode(sChinese, wstrUnicode); // BasicExcel��֧��GB2312�������Խ�����ת��ΪUnicode����

		sheet->Cell(excelIndex,0)->SetWString(wstrUnicode.c_str());
		sheet->Cell(excelIndex,1)->SetString(pinYin.c_str());
		excelIndex++;

		sChinese.clear();
		pinYin.clear();
	}

	int errorLen = errorInfo.size();
	if (errorLen > 0)
	{
		cout << "----------�봦������� " << errorLen << " ������ʾ��ʹ������������----------" << endl << endl;

		for (int errorIndex=0; errorIndex<errorLen; ++errorIndex)
		{
			cout << errorIndex+1 << " " << errorInfo[errorIndex] << endl;
		}

		cout << endl << "----------�봦������� " << errorLen << " ������ʾ��ʹ������������----------" << endl;
		system("pause");
	}
	else
	{
		excel.SaveAs("��ϷͼƬ������.xls");
	}

	errorInfo.clear();

	return 0;
}

string getSpell(const string fileName)
{
	string::size_type pos;
	int i;

	bool find;
	char ch;
	char next;

	string pinYin = "PIC_";
	string strRet;

	int len = fileName.length();
	int index = 0;
	while(fileName[index] != '.' && index < len)
	{
		ch = fileName[index];

		if (isascii(ch))
		{
			pinYin += ch;
			// ֻ������ĸ�����֣������� "." �� "-"
			if (isalpha(ch) || isdigit(ch) || ispunct(ch))
			{
				if (ispunct(ch))
				{
					if (ch != '.' && ch != '_')
					{
						errorInfo.push_back(fileName + "���б�����\"" + ch + "\" �Ƿ�");
					}
				}
			}
			else
			{
				if (ch != ' ')// �ո������ʾ
				{
					errorInfo.push_back(fileName + "���з���\"" + ch + "\" �Ƿ�");
				}
			}
			index += 1;
		}
		else
		{
			find = false;

			next = fileName[index+1];
			strRet = GetPinyinHead(ch, next);

			if ("ERROR" != strRet)
			{
				pinYin += strRet;
				find = true;
			}
			else
			{
				for (i=0; i<PIN_YIN_LENGTH; ++i)
				{
					pos = hanZiData[i].find(fileName.substr(index, 2));
					if (pos != string::npos)
					{
						find = true;
						pinYin += pinYinData[i][pos/2];
						break;
					}
				}
			}


			// ��������������û��ƴ��������ĸ
			if (!find)
			{
				errorInfo.push_back(fileName + "���еĺ���\"" 
					+ fileName[index] + fileName[index+1]
				+ "\"�޷�ʶ����ȷ�ϸú��ֲ������ı�����֮����ϵ�����¬��Ⱥ���Ӹú��ֵ���дƴ��");
			}

			index += 2;
		}
	}

	return pinYin;
}


void gb2312ToUnicode(const string& src, wstring& result)  
{  
	int n = MultiByteToWideChar( CP_ACP, 0, src.c_str(), -1, NULL, 0 );  
	result.resize(n);  
	::MultiByteToWideChar( CP_ACP, 0, src.c_str(), -1, (LPWSTR)result.c_str(), result.length());  
}

void getFiles(string path, vector<string>& files) 
{
	//�ļ����  
	long hFile = 0;  
	//�ļ���Ϣ  
	struct _finddata_t fileinfo;  
	string p;
	if((hFile = _findfirst(p.assign(path).append("/*").c_str(),&fileinfo)) != -1)  
	{
		do{ 
			//�����Ŀ¼,����֮
			//�������,�����б�
			if((fileinfo.attrib & _A_SUBDIR))
			{
				if(strcmp(fileinfo.name,".") != 0 && strcmp(fileinfo.name,"..") != 0)
				{
					getFiles(p.assign(path).append("/").append(fileinfo.name), files);
				}
			}  
			else // ����Ҫ��ȡ��Ŀ¼�µ��ļ�
			{
				files.push_back(p.assign(path).append("/").append(fileinfo.name));
			}  
		}while(_findnext(hFile,&fileinfo) == 0);

		_findclose(hFile);  
	}
}

void checkFileName(const string fileName)
{
	int len = fileName.length();
	if (len <= 4)
	{
		errorInfo.push_back(fileName + "���ļ�������û�к�׺��\".bmp\"");
	}
	else
	{
		if (".BMP" != fileName.substr(len-4, 4))// ���û�ҵ�.BMP��׺
		{
			errorInfo.push_back(fileName + "������bmp�ļ������������߾ܾ�Ϊ��������");
		}

		if (string::npos != fileName.find(" "))
		{
			errorInfo.push_back(fileName + "���пո���ȥ���ļ�����Ŀո�");
		}
	}
}

// �ж�var��ֵ�Ƿ�������֮��
int Between(int var, int lower, int upper)
{
	return (var >= lower) && (var <= upper);
}

// ����һ�����ֵ�ƴ������ĸ
string GetPinyinHead(char HI, char LO)
{
	// ���㺺�ֻ�����, �ֳ� "����ASCII��", ��� "����"
	int val = ((unsigned char)HI << 8) + (unsigned char)LO;

	if (Between(val, 0xB0A1, 0xB0C4)) return "A";
	if (Between(val, 0XB0C5, 0XB2C0)) return "B";
	if (Between(val, 0xB2C1, 0xB4ED)) return "C";
	if (Between(val, 0xB4EE, 0xB6E9)) return "D";
	if (Between(val, 0xB6EA, 0xB7A1)) return "E";
	if (Between(val, 0xB7A2, 0xB8c0)) return "F";
	if (Between(val, 0xB8C1, 0xB9FD)) return "G";
	if (Between(val, 0xB9FE, 0xBBF6)) return "H";
	if (Between(val, 0xBBF7, 0xBFA5)) return "J";
	if (Between(val, 0xBFA6, 0xC0AB)) return "K";
	if (Between(val, 0xC0AC, 0xC2E7)) return "L";
	if (Between(val, 0xC2E8, 0xC4C2)) return "M";
	if (Between(val, 0xC4C3, 0xC5B5)) return "N";
	if (Between(val, 0xC5B6, 0xC5BD)) return "O";
	if (Between(val, 0xC5BE, 0xC6D9)) return "P";
	if (Between(val, 0xC6DA, 0xC8BA)) return "Q";
	if (Between(val, 0xC8BB, 0xC8F5)) return "R";
	if (Between(val, 0xC8F6, 0xCBF0)) return "S";
	if (Between(val, 0xCBFA, 0xCDD9)) return "T";
	if (Between(val, 0xCDDA, 0xCEF3)) return "W";
	if (Between(val, 0xCEF4, 0xD188)) return "X";
	if (Between(val, 0xD1B9, 0xD4D0)) return "Y";
	if (Between(val, 0xD4D1, 0xD7F9)) return "Z";

	return "ERROR";
}


