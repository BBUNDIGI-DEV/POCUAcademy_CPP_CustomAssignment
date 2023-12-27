#define _CRT_SECURE_NO_WARNINGS
#include <iostream>
#include <filesystem>
#include "XlsxParser.h"

namespace fs = std::filesystem;

void coutFileState(eXMLFileState fileState);
int main()
{
	XlsxParser tXlsxParser = XlsxParser();

	string testXMLFilePath = "Chapter7_ResourceFiles\\TestXMLFiles";
	string outputCSVFileDirPath = "Chapter7_ResourceFiles\\GeneratedCSVFiles";
	for (const auto& entry : fs::directory_iterator(testXMLFilePath))
	{
		string filePath = entry.path().u8string();
		cout << "Path : " << filePath << " Parsing Start" << endl;
		tXlsxParser.TryParseXLSVToCSV(filePath, outputCSVFileDirPath);
		auto xmlState = tXlsxParser.GetCurrentState();
		coutFileState(xmlState);
		cout << "=======================================" << endl;
	}

	return 0;
}

void coutFileState(eXMLFileState fileState)
{
	string enumToString;
	switch (fileState)
	{
		case eXMLFileState::loaded:
			enumToString = "loaded";
			break;
		case eXMLFileState::notloaded:
			enumToString = "notloaded";
			break;
		case eXMLFileState::FilePathError:
			enumToString = "FilePathError";
			break;
		case eXMLFileState::FileFormatError:
			enumToString = "FileFormatError";
			break;
	}
	cout << enumToString << endl;
}


