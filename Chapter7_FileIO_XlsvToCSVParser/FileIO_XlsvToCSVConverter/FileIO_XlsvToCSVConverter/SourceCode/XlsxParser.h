#pragma once
#define _CRT_SECURE_NO_WARNINGS

#include <cstring>
#include <fstream>
#include <iostream>
#include <sstream>
using namespace std;

enum class eXMLFileState
{
	notloaded,
	loaded,
	FilePathError,
	FileFormatError,
};

enum class eXMLFileErrorState
{
	notLoaded,
	loadedWell,
	WrongLozengeOpen,
	WrongLozengeClose,
	SheetDataNotFounded,
	ScopeNotClosed,
	WrongXLAddress,
};

class XlsxParser final
{
public:
	XlsxParser();
	~XlsxParser() = default;

	inline eXMLFileState GetCurrentState() const
	{
		return mCurrentState;
	}
	const bool TryParseXLSVToCSV(string filePath, string csvFilePath);
	const void XMLTestReader(string filePath);
private:
	eXMLFileState mCurrentState;
	eXMLFileErrorState mErrorState;

	string CreateOutputCSVFilePath(string inputXMLPath, string outputFolderPath) const;
	bool ConvertXMLAddress(char* address, unsigned int stringSize, unsigned int& cIndex, unsigned int& rIndex) const;

};
