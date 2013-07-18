#pragma once

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

#include <unordered_map>
#include <string>
#include <regex>
#include <boost/tokenizer.hpp>
#include <MsHTML.h>
#include "DOMError.hpp"
#include "iDispatchInvoker.hpp"

using namespace std;

class IHTMLXMLHttpRequestPool
{
public:
  virtual bool Cache(IHTMLXMLHttpRequest* pObj) = 0;
};

class IHTMLXMLHttpRequestWrapper : public iDispatchInvoker<IHTMLXMLHttpRequest>
{
	BEGIN_INVOKER(IHTMLXMLHttpRequestWrapper)

		REG_METHOD(abort)
		REG_METHOD(open)
		REG_METHOD(send)
		REG_METHOD(setRequestHeader)
		REG_METHOD(getResponseHeader)
		REG_METHOD(getAllResponseHeaders)

		//r/w attributes
		REG_ATTR_BASE("timeout")
		REG_ATTR_BASE("withCredentials")

		REG_ATTR_BASE("responseType")

		//read only attributes
		REG_R_ATTR(responseBody)
		REG_R_ATTR(responseText)
		REG_R_ATTR(responseXML)
		REG_R_ATTR(readyState)
		REG_R_ATTR(status)
		REG_R_ATTR(statusText)

		REG_ATTR_BASE("upload")

		//write only attributes
		REG_EVENT(onreadystatechange)

		m_pPool = NULL;

		//escape the only single '. the $2 is the matched string
		//m_js_stringregex = L"((?:(?:'')+)|[^'])(('[^'])|('$))";
		m_js_stringregex = L"(')";
	
		m_jsonp_token = L"callback=";
	CleanUp();

	END_INVOKER

protected:
	IHTMLXMLHttpRequestPool* m_pPool;
	std::wstring m_jsonp_token;

public:
	virtual ~IHTMLXMLHttpRequestWrapper()
	{
	};

	virtual void CleanUp()
	{
		iDispatchInvoker<IHTMLXMLHttpRequest>::CleanUp();

		//clear the request info
		std::swap(m_requestinfo,trans_tuple());
	}

	virtual void AttachPool(IHTMLXMLHttpRequestPool* pPool)
	{
		m_pPool = pPool;
	}

	/**
	Try to cache itself into the object pool before release the memory
	*/
	virtual ULONG STDMETHODCALLTYPE Release() 
	{ 
		ULONG uRet = --m_dwRef; 
		bool bCached = false;

		if(m_dwRef == 0) 
		{
			if(m_pPool)
				bCached = m_pPool->Cache(this);

			if(!bCached)
				delete this; 
			else
				CleanUp();
		}

		DTTRACEMSG_DEBUG(_T("Release(): Ref=%d, cached=%s")
			,uRet
			,bCached ? _T("true") : _T("false")
			);

		return uRet;
	};

	bool IsJSONPRequest()
	{
		bool bRet = false;

		if(m_jsonp_token.length() > 0)
		{
			wstring& strURL = get<url>(m_requestinfo);

			wstring::size_type pos = strURL.find(m_jsonp_token);

			bRet = (pos != wstring::npos);
		}

		return bRet;
	};

	wstring PrepareJSONP(const wstring& strJSON)
	{
		wstring strRet;

		wstring::size_type tokenlength = m_jsonp_token.length();

		strRet.reserve(strJSON.length() + tokenlength + 10);
		strRet = strJSON;

		if(tokenlength > 0)
		{
			wstring& strURL = get<url>(m_requestinfo);

			wstring::size_type pos = strURL.find(m_jsonp_token);

			if(pos != wstring::npos)
			{
				wstring::size_type endPos = strURL.find(L'&',pos);

				if(endPos == wstring::npos)
					endPos = strURL.length();

				strRet.insert(0,L"(");
				strRet.insert(0,strURL.substr(pos+tokenlength,endPos - pos));
				strRet += L");";
			}
		}

		return strRet;
	}

	virtual /* [id][propget] */ HRESULT STDMETHODCALLTYPE get_readyState( /* [out][retval] */  long *p)
	{
		VALID_PARAM_POITNER(p);

		*p = getAttrVal(_T("readyState"));

		return S_OK;
	}

	virtual /* [id][propget] */ HRESULT STDMETHODCALLTYPE get_responseBody( /* [out][retval] */  VARIANT *p)
	{
		VALID_PARAM_POITNER(p);

		_bstr_t bstr = getAttrVal(_T("responseBody"));

		p->vt = VT_BSTR;
		p->bstrVal = bstr.Detach(); 

		return S_OK;
	}

	virtual /* [id][propget] */ HRESULT STDMETHODCALLTYPE get_responseText( 
		/* [out][retval] */  BSTR *p)
	{
		VALID_PARAM_POITNER(p);

		const dictionary& resbody = get<res_body>(m_requestinfo);

		std::wstringstream ostream;

		//join the result to the string
		if(IsJSONPRequest())
			toJSON(ostream,resbody);	
		else
			toString(ostream,resbody);

		//prepare the JSONP reply if the request URL is JSONP format
		_bstr_t bstr = PrepareJSONP(ostream.str()).c_str();
		*p = bstr.Detach(); 

		return S_OK;
	}

	virtual /* [id][propget] */ HRESULT STDMETHODCALLTYPE get_responseXML( 
		/* [out][retval] */  IDispatch **p)
	{
		VALID_PARAM_POITNER(p);

		/*
		JQuery 1.6.2 does't follow the W3C recommended process to check the 
		responseType first
		*/

		/*
		_bstr_t restype = getAttrVal(_T("responseType"));

		if(restype.length( ) > 0 && restype != _bstr_t(_T("document")))
		RaiseException(make_error_code(DOMError_error::INVALID_STATE_ERR));
		*/

		return S_OK;
	}

	/**
	the status attribute must return the result of running these steps:
	1. If the state is UNSENT or OPENED, return 0 and terminate these steps.
	2. If the error flag is set, return 0 and terminate these steps.
	3. Return the HTTP status code.

	http://www.w3.org/Protocols/rfc2616/rfc2616-sec10.html

	Informational 1xx
	Successful		2xx (200 OK)
	Redirection		3xx
	Client Error	4xx (400 Bad Request,408 Request Timeout)
	Server Error	5xx (500 Internal Server Error)
	*/
	virtual /* [id][propget] */ HRESULT STDMETHODCALLTYPE get_status( 
		/* [out][retval] */  long *p)
	{
		VALID_PARAM_POITNER(p);

		*p = getAttrVal(_T("status"));

		//default always return 200
		if(*p == 0)
			*p = 200; 

		return S_OK;

	}

	virtual /* [id][propget] */ HRESULT STDMETHODCALLTYPE get_statusText( 
		/* [out][retval] */  BSTR *p)
	{
		VALID_PARAM_POITNER(p);

		_bstr_t bstr = getAttrVal(_T("statusText"));

		*p = bstr.Detach();

		return S_OK;

	}

	virtual /* [bindable][displaybind][id][propput] */ HRESULT STDMETHODCALLTYPE put_onreadystatechange( 
		/* [in] */ VARIANT v)
	{
		getAttrVal(_T("onreadystatechange")) = v;

		return S_OK;
	}

	virtual /* [bindable][displaybind][id][propget] */ HRESULT STDMETHODCALLTYPE get_onreadystatechange( 
		/* [out][retval] */  VARIANT *p)
	{
		VALID_PARAM_POITNER(p);

		p->pdispVal = getAttrVal(_T("onreadystatechange"));

		if(p->pdispVal)
		{
			p->vt = VT_DISPATCH;
			p->pdispVal->AddRef();
		}
		else
			p->vt = VT_EMPTY;

		return S_OK;
	}

	virtual /* [id] */ HRESULT STDMETHODCALLTYPE abort( void)
	{
		states curState = getCurReadyState();

		if( (curState == HEADERS_RECEIVED) || 
			(curState == LOADING) ||
			(curState == OPENED) //&& m_bSend == true )
			)
		{
			ChangeReadyState(DONE);

			//Fire a progress event named abort.

			//Fire a progress event named loadend.

			//If the upload complete flag is false run these substeps:
			//	1. Set the upload complete flag to true.
			//  2. Fire a progress event named abort on the XMLHttpRequestUpload object.
			//  3. Fire a progress event named loadend on the XMLHttpRequestUpload object.
		}

		ChangeReadyState(UNSENT,false);

		return S_OK;
	}

	virtual /* [id] */ HRESULT STDMETHODCALLTYPE open( 
		/* [in] */  BSTR bstrMethod,
		/* [in] */  BSTR bstrUrl,
		/* [in] */ VARIANT varAsync,
		/* [in][optional] */ VARIANT varUser,
		/* [in][optional] */ VARIANT varPassword)
	{
		//clear the previous request informations first
		CleanUp();

		get<method>(m_requestinfo) = bstrMethod;
		get<url>(m_requestinfo) = bstrUrl;
		get<async>(m_requestinfo) = varAsync.boolVal == VARIANT_TRUE;

		BSTR bstrUser = L"";
		BSTR bstrPWD = L"";

		if(V_VT(&varUser) == VT_BSTR && varUser.bstrVal != NULL)
			bstrUser = varUser.bstrVal;

		get<user>(m_requestinfo) = bstrUser;

		if(V_VT(&varPassword) == VT_BSTR && varUser.bstrVal != NULL)
			bstrPWD = varPassword.bstrVal;

		get<pwd>(m_requestinfo) = bstrPWD;

		ChangeReadyState(OPENED);

		return S_OK;
	}

	/**
	This is the helper function to process the send request. It is invoked by the send() after 
	extracting all the request information. the child class should implement this funciton 
	to handle all the URL request. 

	\return bool	true means the request is processed completely. the readyState will
	be updated as DONE by the send().otherwise, the caller need wait for the onreadystatechange
	event when the request is done for the Async request.
	*/
	virtual bool ProcessSendRequest()
	{
		return true;
	};

	/**
	This method takes one optional parameter, which is the requestBody to use. 
	The acceptable VARIANT input types are BSTR, SAFEARRAY of UI1 (unsigned bytes), 
	IDispatch to an XML Document Object Model (DOM) object, and IStream *. You can 
	use only chunked encoding (for sending) when sending IStream * input types. The 
	component automatically sets the Content-Length header for all but IStream * 
	input types.

	in W3C spec, there are only 3 formats call and recommend: the argument is ignored 
	if request method is GET or HEAD.

	void send();
	void send(ArrayBuffer data);
	void send(Blob data);
	void send(Document data);
	void send(DOMString? data);
	void send(FormData data);
	*/
	virtual /* [id] */ HRESULT STDMETHODCALLTYPE send( 
		/* [in][optional] */ VARIANT varBody)
	{
		if(getCurReadyState() != OPENED)
			RaiseException(make_error_code(DOMError_error::INVALID_STATE_ERR),
			_T("Fail to send the request because the current readyState shoould be OPENED")
			);

		wstring strBody;

		if(V_VT(&varBody) == VT_BSTR)
			strBody = V_BSTR(&varBody);

		//parse the parameters 'xxx=xxx&xxx=xxx&...'
		ParseURLParam(get<req_body>(m_requestinfo), strBody);

		if(get<async>(m_requestinfo))	//async process
		{
			//The state does not change. The event is dispatched for historical reasons.
			ChangeReadyState(SEND);	

			//Fire a progress event named loadstart.

			//If the upload complete flag is unset, fire a progress event named loadstart 
			//on the XMLHttpRequestUpload object.
		}
		else
		{
			//sync process
		}

		//default implementation is always done for this action. The detail
		//process should be implemented in the derived class
		if(ProcessSendRequest())
			ChangeReadyState(DONE);	

		return S_OK;
	}

	virtual /* [id] */ HRESULT STDMETHODCALLTYPE getAllResponseHeaders( 
		/* [out][retval] */  BSTR * p)
	{
		VALID_PARAM_POITNER(p);

		std::wstringstream ostream;

		toString(ostream,get<res_header>(m_requestinfo));

		_bstr_t bstr(ostream.str().c_str());
		*p = bstr.Detach();

		return S_OK;
	}

	virtual /* [id] */ HRESULT STDMETHODCALLTYPE getResponseHeader( 
		/* [in] */  BSTR bstrHeader,
		/* [out][retval] */  BSTR * p)
	{
		VALID_PARAM_POITNER(p);

		//need re-allocate BSTR since it is out parameter
		_bstr_t bstr = get<res_header>(m_requestinfo)[bstrHeader].c_str();
		*p = bstr.Detach();
		return S_OK;
	}

	virtual /* [id] */ HRESULT STDMETHODCALLTYPE setRequestHeader( 
		/* [in] */  BSTR bstrHeader,
		/* [in] */  BSTR bstrValue)
	{
		get<req_header>(m_requestinfo)[bstrHeader] = bstrValue;
		return S_OK;
	}

protected:
	/*const std::string XMLHttpRequestResponseType[] =  
	{
	"",
	"arraybuffer",
	"blob",
	"document",
	"json",
	"text"
	};*/

	enum states
	{
		SEND = -1, /** It is not the standard state */

		UNSENT = 0,
		OPENED = 1,
		HEADERS_RECEIVED = 2,
		LOADING = 3,
		DONE = 4
	};

	inline states getCurReadyState()
	{
		return (states)(getAttrVal(_T("readyState")).intVal);
	};

	inline void ChangeReadyState(states enState, bool bFireEvent=true)
	{
		if(enState != SEND)
			getAttrVal(_T("readyState")) = enState;

		switch(enState)
		{
		case OPENED:
			break;

		case SEND:
			break;

		default:
			break;
		}

		//fire the event
		if(bFireEvent)
			onreadystatechange();
	}

	wstring getRequestBody(wstring const & pkey)
	{
		const dictionary& reqParams = get<req_body>(m_requestinfo);
		wstring retVal;

		dictionary::const_iterator iter = reqParams.find(pkey);
		if(iter != reqParams.end())
			retVal = iter->second;

		return retVal;
	};

	inline void setResBody(wstring const& key, wstring const & val)
	{
		get<res_body>(m_requestinfo)[key] = val;
	};

protected:
	wregex m_js_stringregex;
	
	typedef unordered_map<wstring,wstring> dictionary;

	/** index for the tuple */
	enum {method=0,url,async,user,pwd, req_header,req_body,res_header,res_body};

	/** store the request/response detail information */
	typedef boost::fusion::tuple<wstring,wstring,bool,wstring,wstring,dictionary,dictionary,dictionary,dictionary> trans_tuple;
	trans_tuple m_requestinfo;

	inline std::wstringstream& toJSON(std::wstringstream& ostream, dictionary const& map)
	{
		ostream << "{";

		std::for_each(map.begin(),map.end(),[&](dictionary::value_type const& item){
		
			//check whether the value is array or oject
			switch(item.second[0])
			{
			case L'[':
			case L'{':
				ostream << item.first << ":" <<  item.second << ",";
				break;

			default:
				ostream << item.first << ":'" << regex_replace(item.second,m_js_stringregex,(wchar_t*)L"\\$1") << "',";
				break;
			}
		});

		if(ostream.tellp() > 1)
			ostream.seekp(-1,ios_base::cur);

		ostream << "}";

		return ostream;
	};

	inline std::wstringstream& toString(std::wstringstream& ostream,dictionary const & map, wstring::const_pointer pDelimiter = L"\n")
	{
		std::for_each(map.begin(),map.end(),[&ostream,&pDelimiter](dictionary::value_type const& item){
			ostream << item.first << L"=" << item.second << pDelimiter;
		});

		return ostream;
	};

	//parse the parameters 'xxx=xxx&xxx=xxx&...'
	size_t ParseURLParam(dictionary& keyvalMap,wstring& strVal)
	{
		size_t count = 0;

		typedef boost::tokenizer<boost::char_separator<wstring::value_type>,wstring::const_iterator,wstring> req_token;
		req_token tok(strVal,boost::char_separator<wstring::value_type>(L"&"));

		for_each(tok.begin(),tok.end(),[&](req_token::const_reference token)
		{
			req_token items(token,boost::char_separator<wstring::value_type>(L"="));

			req_token::iterator iter = items.begin();
			req_token::iterator iterEnd = items.end();

			if(iter != iterEnd)
			{
				wstring strval;
				wstring strkey = *iter;
				iter++;

				if(iter != iterEnd)
					strval = *iter;

				keyvalMap[strkey] = strval;

				count++;
			}
		});

		return count;
	};

	/** invoke the registered call back function following the W3C semantics */
	inline void onreadystatechange()
	{
		HRESULT hr = S_OK;
		IDispatch* pDisp = NULL;

		_variant_t& var = getAttrVal(_T("onreadystatechange"));

		if((V_VT(&var) & VT_DISPATCH) == VT_DISPATCH)
			pDisp = V_DISPATCH(&var);

		if(pDisp!= NULL)
		{
			// Invoke method with "this" pointer
			DISPPARAMS dispparams;
			DISPID putid = DISPID_THIS;
			var.vt = VT_DISPATCH;
			var.pdispVal = pDisp;

			dispparams.rgvarg = &var;
			dispparams.cArgs = 1;

			dispparams.rgdispidNamedArgs = &putid;
			dispparams.cNamedArgs = 1;

			hr = pDisp->Invoke(DISPID_VALUE,IID_NULL,LOCALE_USER_DEFAULT,DISPATCH_METHOD,&dispparams,NULL,NULL,NULL);

			DTTRACEMSG_DEBUG(_T("Invoke the onreadystatechange(): hr=0X%X (%s)")
				,hr
				,DTHRESULT_ERRMSG(hr).c_str()
				);
		}
	}
};
