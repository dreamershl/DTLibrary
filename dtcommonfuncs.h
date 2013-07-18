/**
 * (C) Copyright 2013 Dreamer
 *
 * this file to You under the Apache License, Version 2.0
 * (the "License"); you may not use this file except in compliance with
 * the License.  You may obtain a copy of the License at
 *
 *     http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */

/**
* Define the common useful functions in namespace DT
*
*/

#ifndef __DTCOMMONFUNCS_H__
#define __DTCOMMONFUNCS_H__

#include "WinInet.h"
#include "DispEx.h"

#include <string>

namespace DT
{
  /**
	* Return the specified address module handle. DON'T FREE the return 
	* handle. 
	*/
	inline HMODULE GetAddrModule(void* pAddr)
	{
		HMODULE hModule = NULL;

		GetModuleHandleEx(
			GET_MODULE_HANDLE_EX_FLAG_UNCHANGED_REFCOUNT|
			GET_MODULE_HANDLE_EX_FLAG_FROM_ADDRESS,
			(LPCTSTR) pAddr,
			&hModule
			);

		return hModule;
	}


	/**
	* WinInet Global functions
	*/
	inline bool IsAvailableOffline(LPCTSTR pszURL)
	{
		HRESULT hr;
		DWORD   dwUsesNet, dwCached;
		DWORD   dwSize;

		if (!pszURL)
			return false;

		_bstr_t bstrURL = pszURL;
		LPCWSTR pwszURL = bstrURL;

		// First, let URL monikers check the protocol scheme.
		hr = CoInternetQueryInfo(pwszURL, 
			QUERY_USES_NETWORK,
			0,
			&dwUsesNet,
			sizeof(dwUsesNet),
			&dwSize,
			0);

		if (FAILED(hr) || !dwUsesNet)
			return true;

		// Then let URL monikers peek in the cache.
		hr = CoInternetQueryInfo(pwszURL, 
			QUERY_IS_CACHED_OR_MAPPED, 
			0,
			&dwCached, 
			sizeof(dwCached), 
			&dwSize, 
			0);

		if (FAILED(hr))
			return false;

		return true;
	}

	inline void SetWinInetGlobalOffline(bool fGoOffline)
	{
		INTERNET_CONNECTED_INFO ci;

		memset(&ci, 0, sizeof(ci));
		if(fGoOffline) 
		{
			ci.dwConnectedState = INTERNET_STATE_DISCONNECTED_BY_USER;
			ci.dwFlags = ISO_FORCE_DISCONNECTED;
		} 
		else 
		{
			ci.dwConnectedState = INTERNET_STATE_CONNECTED;
		}

		InternetSetOption(NULL, INTERNET_OPTION_CONNECTED_STATE, &ci, sizeof(ci));
	}

	// Returns true if the global state is offline. Otherwise, false. 
	inline bool IsWinInetGlobalOffline()
	{
		DWORD dwState = 0; 
		DWORD dwSize = sizeof(DWORD);
		bool fRet = false;

		if(InternetQueryOption(NULL, INTERNET_OPTION_CONNECTED_STATE, &dwState, &dwSize))
		{
			if(dwState & INTERNET_STATE_DISCONNECTED_BY_USER)
				fRet = true;
		}

		return fRet; 
	}


	/***
	*	MSDHTML Helper functions
	*
	*/

	inline HRESULT ExecJSScript(IHTMLDocument2* pDoc,LPCTSTR szScript)
	{
		if(pDoc == NULL)
			return E_POINTER;

		HRESULT hr = S_OK;
		CComBSTR oBSTR,oBSTRLanguage;

		oBSTR = szScript;
		oBSTRLanguage = "javascript";

		COleVariant ret;
		CComPtr<IHTMLWindow2> oWindow2;
		hr = pDoc->get_parentWindow(&oWindow2);
		ASSERT(SUCCEEDED(hr));

		hr = oWindow2->execScript(oBSTR, oBSTRLanguage, &ret);

		return hr;
	}

	inline HRESULT ExecJSFunc(IHTMLDocument2* pDoc, LPCTSTR szFunc, const CStringArray& paramArray,CComVariant* pVarResult=NULL)
	{
		if(pDoc == NULL)
			return E_POINTER;

		HRESULT hr;

		CComPtr<IDispatch> spScript;
		hr = pDoc->get_Script(&spScript);

		if(SUCCEEDED(hr))
		{
			CComBSTR bstr(szFunc);
			DISPID dispid = NULL;

			hr = spScript->GetIDsOfNames(IID_NULL,&bstr,1,
				LOCALE_SYSTEM_DEFAULT,&dispid);
			if(SUCCEEDED(hr))
			{
				int arraySize = (int)paramArray.GetSize();

				DISPPARAMS dispparams;
				memset(&dispparams, 0, sizeof dispparams);
				dispparams.cArgs = (UINT)arraySize;
				dispparams.rgvarg = new VARIANT[dispparams.cArgs];

				for( int i = 0; i < arraySize; i++)
				{
					dispparams.rgvarg[i].bstrVal = A2WBSTR(paramArray.GetAt(arraySize - 1 - i)); // back reading
					dispparams.rgvarg[i].vt = VT_BSTR;
				}

				EXCEPINFO excepInfo;
				memset(&excepInfo, 0, sizeof excepInfo);

				CComVariant vaResult;
				UINT nArgErr = (UINT)-1;  // initialize to invalid arg

				hr = spScript->Invoke(dispid,IID_NULL,0,
					DISPATCH_METHOD,&dispparams,&vaResult,&excepInfo,&nArgErr);

				for(int i = 0; i < arraySize; i++)
					::SysFreeString(dispparams.rgvarg[i].bstrVal);

				delete [] dispparams.rgvarg;

				if(SUCCEEDED(hr))
				{
					if(pVarResult)
						*pVarResult = vaResult;
				}
			}
		}

		return hr;
	}

	inline HRESULT LoadJSScript(IHTMLDocument2* pDoc,LPCTSTR szUrl, bool bDefer = true, bool bLast=true)
	{
		CString strScript;

		strScript.Format("var script = document.createElement(\"script\");"
			"script.src = '%s'; script.type=\"text/Javascript\";"
			"%s"
			"var parent = document.getElementsByTagName(\"head\")[0] || document.body;"
			,szUrl
			,bDefer ? _T("script.defer=true;") : _T(""));

		if(bLast)
			strScript += _T("parent.appendChild( script );");
		else
			strScript += _T("parent.insertBefore( script, parent.firstChild);");

		return ExecJSScript(pDoc, strScript);	
	}

	inline HRESULT LoadResJSScript(IHTMLDocument2* pDoc, LPCTSTR szExeName,int nResID,bool bDefer = true,bool bLast = true)
	{
		CString strURL;

		strURL.Format(_T("res://%s.exe/#%d"),szExeName, nResID);

		return LoadJSScript(pDoc,strURL,bDefer,bLast);
	}

	inline HRESULT LoadResJSScript(IHTMLDocument2* pDoc, LPCTSTR szExeName,LPCTSTR szResID, bool bDefer = true,bool bLast = true)
	{
		CString strURL;

		strURL.Format(_T("res://%s.exe/%s"),szExeName, szResID);

		return LoadJSScript(pDoc,strURL,bDefer,bLast);
	}

	inline HRESULT RegisterJSObject(LPCTSTR szObjName,IDispatch* pObj, IHTMLWindow2* pHTMLWin)
	{
		CComQIPtr<IDispatchEx> pWinEx(pHTMLWin);

		if (pWinEx == NULL) 
			return DISP_E_UNKNOWNINTERFACE;

		HRESULT hr;
		CComBSTR objName = szObjName;
		DISPID dispid; 
		//register the object name first
		hr = pWinEx->GetDispID(objName, fdexNameEnsure, &dispid);

		if (SUCCEEDED(hr)) 
		{
			DISPID namedArgs[] = {DISPID_PROPERTYPUT};
			DISPPARAMS params;
			
			VARIANT varVal;
			varVal.pdispVal = pObj;
			varVal.vt = VT_DISPATCH;
			
			params.rgvarg = &varVal;
			params.rgdispidNamedArgs = namedArgs;
			params.cArgs = 1;
			params.cNamedArgs = 1;
			
			hr = pWinEx->InvokeEx(dispid, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYPUT, &params, NULL, NULL, NULL); 
		}

		return hr;
	}

	inline HRESULT RegisterJSObject(LPCTSTR szObjName,IDispatch* pObj, IHTMLDocument2* pDoc)
	{
		CComPtr<IHTMLWindow2> oWindow2;
		HRESULT hr = pDoc->get_parentWindow(&oWindow2);
		
		if(SUCCEEDED(hr))
			hr = RegisterJSObject(szObjName,pObj,(IHTMLWindow2*)oWindow2);

		return hr;
	};

	inline void preventDefault(IHTMLEventObj* pEvent)
	{
		if(pEvent)
		{
			_variant_t varVal;
			varVal = VARIANT_FALSE;
			//prevent default event handler
			pEvent->put_returnValue(varVal);
		}
	}

	inline std::string GetGUIDName(REFGUID iid)
	{
		OLECHAR szBuf[100] = {0};
		::StringFromGUID2(iid, szBuf,sizeof(szBuf));

		CString strSubKey;
		strSubKey.Format(_T("Interface\\%ls"),szBuf); 

		CRegKey reg;
		bool bRawID = true;

		TCHAR szVal[200] = {0};
		ULONG ulSize = sizeof(szVal);

		if(reg.Open(HKEY_CLASSES_ROOT,strSubKey,KEY_READ) != ERROR_SUCCESS)
		{
			strSubKey.Format(_T("CLSID\\%ls"),szBuf);
			reg.Open(HKEY_CLASSES_ROOT,strSubKey,KEY_READ);
		}

		if(reg.m_hKey != NULL)
			bRawID = (reg.QueryStringValue(NULL,szVal,&ulSize) != ERROR_SUCCESS);

		if(!bRawID)
			strSubKey.Format(_T("%s - %ls"), szVal, szBuf);
		else
			strSubKey.Format(_T("%ls"), szBuf);

		return std::string(strSubKey);
	}

	inline bool IsUTF16(LPCWSTR szBuf, int nSize)
	{
		return (::WideCharToMultiByte(
			CP_UTF8,            // convert to UTF-8
			0,                  // default flags
			szBuf,       // source UTF-16 string
			nSize,     // source string length, in wchar_t's,
			NULL,               // unused - no conversion required in this step
			0,                  // request buffer size
			NULL, NULL          // unused
			) != 0);
	}



	/**
	* Converts a string from UTF-16 to UTF-8.
	* @param nSize in,out it specified the szBuf char length (not byte size) for the converstion. if
	*						convertion is failed, it will be <=0 for the error code.
	* 
	* @return string will return empty string only if the input string is UTF16 but can't convert to 
	*				UTF8 successfully. the param nSize will be 
	*				-1 => the input string is not UTF16, the return std::string will be assigned the input chars
	*				-2 => the input string is UTF16, but convertion fail. call GetLastError() to get the detail
	*/
	inline std::string toUTF8(LPCWSTR szBuf, int& nSize)
	{
		std::string strRet;
		//
		// Get length (in chars) of resulting UTF-8 string
		//
		int nUTF8Size = ::WideCharToMultiByte(
			CP_UTF8,            // convert to UTF-8
			0,                  // default flags
			szBuf,				// source UTF-16 string
			nSize,				// source string length, in wchar_t's,
			NULL,               // unused - no conversion required in this step
			0,                  // request buffer size
			NULL, NULL          // unused
			);

		if (nUTF8Size != 0)
		{
			//
			// Allocate destination buffer for UTF-8 string
			//
			strRet.resize(nUTF8Size);

			//
			// Do the conversion from UTF-16 to UTF-8
			//
			if ( ! ::WideCharToMultiByte(
				CP_UTF8,                // convert to UTF-8
				0,                      // default flags
				szBuf,					// source UTF-16 string
				nSize,					// source string length, in wchar_t's,
				(LPSTR)strRet.data(),          // destination buffer
				(int)(strRet.length()),        // destination buffer size, in chars
				NULL, NULL              // unused
				) )
				nSize = -2;
		}
		else 
		{
			strRet.assign((LPCTSTR)szBuf,nSize);
			nSize = -1;
		}

		return strRet;
	}

	inline bool IsUTF8(LPCSTR szBuf,  int nSize )
	{
		return (MultiByteToWideChar(CP_UTF8,MB_ERR_INVALID_CHARS,szBuf, nSize, NULL,0) != 0);
	}

	/**
	* Use the "CreateDircectory() to create the subfolders. The last part is always treated as the file name
	* So if need create all the components, the input path should be like "xx1\xx2\xx3\"
	* 
	* @return int 0 means success, otherwise it is the error code of GetLastError() if the API CreateDirectory() fail. 
	*/
	inline int CreateFolderTree(const std::string & strPath)
	{
		int nError = 0;
		bool bFlag = true;
		std::string::size_type nPos = 0;
		std::string strFolder;

		do
		{
			 nPos = strPath.find(_T('\\'),nPos);

			if(nPos != std::string::npos)
			{
				strFolder = strPath.substr(0,nPos);
				nPos++;
			}
			else
			{
				nError = 0;
				break;
			}

			if( ::CreateDirectory(strFolder.c_str(),NULL) == FALSE)
			{
				nError = GetLastError();
				bFlag = (nError == ERROR_ALREADY_EXISTS);
			}

		}while(bFlag);

		return nError;
	}

	// Return the protocol associated with the specified URL
	inline INTERNET_SCHEME GetScheme(LPCTSTR szURL)
	{
		URL_COMPONENTS urlComponents;
		DWORD dwFlags = 0;
		INTERNET_SCHEME nScheme = INTERNET_SCHEME_UNKNOWN;

		ZeroMemory((PVOID)&urlComponents, sizeof(URL_COMPONENTS));
		urlComponents.dwStructSize = sizeof(URL_COMPONENTS);

		if (szURL)
		{
			if (InternetCrackUrl(szURL, 0, dwFlags, &urlComponents))
			{
				nScheme = urlComponents.nScheme;
			}
		}

		return nScheme;
	}

	/**
	* It is to set the caller thread COM error object through the SetErrorInfo(). After set this, the 
	* caller function MUST NOT invoke any COM function. Otherwise, the set error info will be overwrited. 
	*/
	inline HRESULT SetComErrorInfo(REFGUID srcGUID,LPCTSTR szErrMsg, LPCTSTR szSource = NULL,LPCTSTR szHelpFile=NULL, DWORD dwHelpContext = 0 )
	{
		ICreateErrorInfo *perrinfo;
		HRESULT hr;

		_bstr_t strSrc, strMsg, strHelp;

		strSrc = szSource;
		strMsg = szErrMsg;
		strHelp = szHelpFile;

		hr = CreateErrorInfo(&perrinfo);

		if(SUCCEEDED(hr))
		{
			perrinfo->SetGUID(srcGUID);
			perrinfo->SetSource(strSrc);
			perrinfo->SetDescription(strMsg);
			perrinfo->SetHelpFile(strHelp);
			perrinfo->SetHelpContext(dwHelpContext);

			IErrorInfo* pInfo;

			hr = perrinfo->QueryInterface(IID_IErrorInfo, (LPVOID FAR*) &pInfo);

			if (SUCCEEDED(hr))
			{
				SetErrorInfo(0, pInfo);
				pInfo->Release();
			}

			perrinfo->Release();
		}

		return hr;
	}

	/**
	* It is to set the caller thread COM error object through the SetErrorInfo(). After set this, the 
	* caller function MUST NOT invoke any COM function. Otherwise, the set error info will be overwrited. 
	*/
	inline HRESULT SetComErrorInfo(LPCTSTR szSource,LPCTSTR szFmt,...)
	{
		va_list args;
		TCHAR szMsg[100*1024]; //allocate the memory from stack, faster
	
		va_start(args, szFmt);
		
		int nLength = sizeof(szMsg) - 1;
		TCHAR* pBuf = szMsg;
		int nResult;

		szMsg[nLength] = 0;
		nResult = _vsnprintf(pBuf, nLength ,szFmt, args);

		if(nResult < 0 )
		{
			pBuf = _T("Can't store the error detail because the message is bigger than the buffer, 100K");
		}

		HRESULT hr = SetComErrorInfo(GUID_NULL, pBuf,szSource);

		va_end(args);

		return hr;
	}
}

#endif
