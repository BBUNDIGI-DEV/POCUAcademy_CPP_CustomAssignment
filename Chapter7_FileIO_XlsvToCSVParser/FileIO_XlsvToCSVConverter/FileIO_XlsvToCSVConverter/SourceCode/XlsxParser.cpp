#include "XlsxParser.h"
#include <cassert>
#include <filesystem>

#define DATA_BLOCK_BUFFER_SIZE 256

XlsxParser::XlsxParser()
	:mCurrentState(eXMLFileState::notloaded), mErrorState(eXMLFileErrorState::notLoaded)
{

}

const bool XlsxParser::TryParseXLSVToCSV(string xmlFilePath, string csvFilePath)
{
	ifstream xlsxFin;
	xlsxFin.open(xmlFilePath);
	
	if (!xlsxFin.is_open())
	{
		mCurrentState = eXMLFileState::FileFormatError;
		return false;
	}
	string readLine;
	char dataBlock[DATA_BLOCK_BUFFER_SIZE];
	char readChar;

	bool isSheetDataExist = false;
	unsigned int startColumnDimension = 0;
	unsigned int startRowDimension = 0;
	unsigned int lastColumnDimension = 0;
	unsigned int lastRowDimension = 0;
	unsigned int dimensionWidth = 0;
	unsigned int dimensionHeight = 0;

	getline(xlsxFin, readLine);

	while (!xlsxFin.eof())
	{
		readChar = xlsxFin.get();
		if (readChar != '<')
		{
			mErrorState = eXMLFileErrorState::WrongLozengeOpen;
			mCurrentState = eXMLFileState::FileFormatError;
			return false;
		}

		getline(xlsxFin, readLine, '>');
		stringstream sStream(readLine);


		sStream >> dataBlock;
		if (strcmp(dataBlock, "sheetData") == 0)
		{
			cout << "SheetData Get Chapterd" << endl;
			isSheetDataExist = true;
			break;
		}
		else if (strcmp(dataBlock, "dimension") == 0)
		{
			sStream >> dataBlock;
			char* token = strtok(dataBlock, "=");
			assert(strcmp(dataBlock, "ref") == 0);
			token = strtok(NULL, "\"");
			
			token = strtok(token, ":");
			ConvertXMLAddress(token, strlen(token), startColumnDimension, startRowDimension);

			token = strtok(NULL, ":");
			ConvertXMLAddress(token, strlen(token), lastColumnDimension, lastRowDimension);

			dimensionWidth = lastColumnDimension - startColumnDimension + 1;
			dimensionHeight = lastRowDimension - startRowDimension + 1;
		}
	}

	if (!isSheetDataExist)
	{
		mErrorState = eXMLFileErrorState::SheetDataNotFounded;
		mCurrentState = eXMLFileState::FileFormatError;
		return false;
	}

	//Prepare SheetParsing
	bool isHappyPass = false;
	ofstream convertedCSVFile;
	convertedCSVFile.open(CreateOutputCSVFilePath(xmlFilePath, csvFilePath));
	string prevReadDataBlock;
	string flushString;
	unsigned int currentColumnIndex = 0;
	unsigned int currentRowIndex = 0;
	unsigned int prevColumnIndex = -1;
	unsigned int prevRowIndex = -1;
	unsigned int blockCount = 0;
	bool isInitialWrite = true;
	bool isInitialLine = true;

	while (!xlsxFin.eof())
	{
		getline(xlsxFin, readLine, '>');
		stringstream sStream(readLine);
		char singleChar;

		singleChar = sStream.get();
		assert(singleChar == '<');

		sStream >> dataBlock;
		if (strcmp(dataBlock, "row")==0)
		{
			prevReadDataBlock = "row";
			blockCount++;
			sStream >> dataBlock;
			char* token = strtok(dataBlock, "=");
			if (strcmp(token, "r") == 0)
			{
				token = strtok(NULL, "\"");
				currentRowIndex = stoi(token) - 1;
			}
			else if (strcmp(token, "spans") == 0)
			{
				token = strtok(NULL, "\"");
				//token now have spans data but this will not used
			}
			
			if (isInitialWrite)
			{
				isInitialWrite = false;
				assert(startRowDimension == currentRowIndex);
			}
			else
			{
				int rowOffset = currentRowIndex - prevRowIndex;
				if (rowOffset > 1)
				{
					for (unsigned int i = 0; i < rowOffset; i++)
					{
						for (unsigned int j = 0; j < dimensionWidth; j++)
						{
							convertedCSVFile << ',';
						}
						convertedCSVFile << endl;
					}
				}
			}
			isInitialLine = true;
			prevRowIndex = currentRowIndex;
		}
		else if (strcmp(dataBlock, "c") == 0)
		{
			assert(strcmp(prevReadDataBlock.c_str(), "row") == 0 ||
			strcmp(prevReadDataBlock.c_str(), "\c") == 0);
			prevReadDataBlock = "c";
			blockCount++;
			sStream >> dataBlock;
			char* token = strtok(dataBlock, "=");
			if ((strcmp(token, "r") == 0))
			{
				token = strtok(NULL, "\"");
				unsigned int stringSize;
				ConvertXMLAddress(token,strlen(token), currentColumnIndex, currentRowIndex);
			}
		}
		else if (strcmp(dataBlock, "v") == 0)
		{
			assert(strcmp(prevReadDataBlock.c_str(), "c") == 0);
			getline(xlsxFin, readLine, '<');
			//Now readline Is Contents
			if (isInitialLine)
			{
				isInitialLine = false;

				unsigned int columnOffset = currentColumnIndex - startColumnDimension;
				if (columnOffset > 0)
				{
					for (unsigned int i = 0; i < columnOffset; i++)
					{
						convertedCSVFile << ',';
					}
				}
			}
			else
			{
				unsigned int columnOffset = currentColumnIndex - prevColumnIndex;
				if (columnOffset > 1)
				{
					for (unsigned int i = 0; i < columnOffset - 1; i++)
					{
						convertedCSVFile << ',';
					}
				}
				convertedCSVFile << ',';
			}
			convertedCSVFile << readLine;

			prevColumnIndex = currentColumnIndex;

			getline(xlsxFin, flushString, '>');
			assert(strcmp(flushString.c_str(), "/v") == 0);
			//Flushing Until /v Met;
		}
		else if (strcmp(dataBlock, "/row") == 0)
		{
			unsigned int columnOffset = lastColumnDimension - currentColumnIndex;
			if (columnOffset > 1)
			{
				for (int i = 0; i < columnOffset; i++)
				{
					convertedCSVFile << ',';
				}
			}

			convertedCSVFile << endl;
			blockCount--;
		}
		else if (strcmp(dataBlock, "/c") == 0)
		{
			blockCount--;
		}
		else if (strcmp(dataBlock, "/sheetData"))
		{
			break;
		}
	}
	xlsxFin.close();
	convertedCSVFile.close();
	mCurrentState = eXMLFileState::loaded;
	return false;
}

const void XlsxParser::XMLTestReader(string filePath)
{
	ifstream xlsFin;
	xlsFin.open(filePath,ios_base::in);



	if (!xlsFin.is_open())
	{
		cout << "File Is Not Opened" << endl;
		return;
	}

	string word;

	while (!xlsFin.eof())
	{
		xlsFin >> word;
		cout << word << endl;
		float trash;
		cin >> trash;
	}
	return void();
}

string XlsxParser::CreateOutputCSVFilePath(string inputXMLPath, string outputFolderPath) const
{
	string outputString = outputFolderPath;
	const char* inputPathCStr = inputXMLPath.c_str();

	unsigned int inputPathLen = strlen(inputPathCStr);
	char xmlFileName[64] = {'\0'};
	for (int i = inputPathLen; i >= 0; i--)
	{
		if (inputPathCStr[i] == '\\')
		{
			strcpy(xmlFileName, &inputPathCStr[i]);
			break;
		}
	}
	assert(xmlFileName != nullptr);
	size_t xmlFileNameLength = strlen(xmlFileName);
	for (unsigned int i = xmlFileNameLength - 1; i > xmlFileNameLength - 1 - 3; i--)
	{
		int backwardIndex = xmlFileNameLength -1 - i;

		switch (backwardIndex)
		{
		case 0:
			xmlFileName[i] = 'v';
			break;
		case 1:
			xmlFileName[i] = 's';
			break;
		case 2:
			xmlFileName[i] = 'c';
			break;
		default:
			break;
		}
	}

	outputString.append(xmlFileName);
	return outputString;
}

bool XlsxParser::ConvertXMLAddress(char* address, unsigned int stringSize, unsigned int& cIndex, unsigned int& rIndex) const
{
	unsigned int numberStartIndex = UINT_MAX;
	for (unsigned int i = 0; i < stringSize; i++)
	{
		if (address[i] >= '1' && address[i] <= '9')
		{
			numberStartIndex = i;
			break;
		}
	}

	if (numberStartIndex == UINT_MAX)
	{
		return false;
	}

	const int alphabetCount = 26;
	cIndex = 0;
	for (unsigned int i = 0; i < numberStartIndex; i++)
	{
		int manipuler = static_cast<int>(pow(alphabetCount, numberStartIndex - 1 - i));
		manipuler = manipuler * (address[i] - 'A' + 1);
		cIndex += manipuler;
	}

	rIndex = 0;
	for (unsigned int i = numberStartIndex; i < stringSize; i++)
	{
		int manipuler = static_cast<int>(pow(10, stringSize - 1 - i));
		manipuler = manipuler * (address[i] - '0');
		rIndex += manipuler;

	}
	rIndex--;
	return true;
}
