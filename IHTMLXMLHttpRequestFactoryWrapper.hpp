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


#include <forward_list>
#include <atlcomcli.h>
#include <MsHTML.h>
#include "iDispatchInvoker.hpp"
#include "IHTMLXMLHttpRequestWrapper.h"

/**
The caller can specify the IHTMLXMLHttpRequest class through the template parameter. The class should be child class 
of IHTMLXMLHttpRequestWrapper to provide the interface AttachFactory(). This factor uses the pool to cache the created 
IHTMLXMLHttpRequest objects. When receives the second request, it will check avaible object in the pool before creating 
a new object
*/
template<typename requestclass=IHTMLXMLHttpRequestWrapper, int cacheLmt = 100> 
class IHTMLXMLHttpRequestFactoryWrapper :
  public iDispatchInvoker<IHTMLXMLHttpRequestFactory>, public IHTMLXMLHttpRequestPool
{
protected:
	std::forward_list<IHTMLXMLHttpRequest*> m_objpool;

public:
	BEGIN_INVOKER(IHTMLXMLHttpRequestFactoryWrapper)
		REG_METHOD_DEFAULT(_T("create"),create)
	END_INVOKER

	virtual ~IHTMLXMLHttpRequestFactoryWrapper(void)
	{
		//release the stored objects in the pool
		for_each(m_objpool.begin(),m_objpool.end(), [](IHTMLXMLHttpRequest* pObj)
		{
			//stop the cache process
			((IHTMLXMLHttpRequestWrapper*)pObj)->AttachPool(NULL);
			pObj->Release();
		}); 
	}

	virtual /* [id] */ HRESULT STDMETHODCALLTYPE create( /* [out][retval] */IHTMLXMLHttpRequest **pIHTMLXMLHttpRequest)
	{
		VALID_PARAM_POITNER(pIHTMLXMLHttpRequest);

		if(m_objpool.empty())
		{
			*pIHTMLXMLHttpRequest = new requestclass();
			((IHTMLXMLHttpRequestWrapper*)*pIHTMLXMLHttpRequest)->AttachPool(this);
		}
		else
		{
			*pIHTMLXMLHttpRequest = m_objpool.front();
			m_objpool.pop_front();
		}

		return S_OK;
	}

	/**
	Based on the cache limit to determine whether cache the IHTMLXMLHttpRequest object. return true means the 
	object is stored in the pool for the next time usage. Then that object should not be released anymore. 
	This function is called in the IHTMLXMLHttpRequest::Release(). 

	\return bool true: object is cached. that object should not be released anymore
	*/
	virtual bool Cache(IHTMLXMLHttpRequest* pObj)
	{
		bool bRet = false;

		if(std::distance(m_objpool.begin(), m_objpool.end()) < cacheLmt)
		{
			pObj->AddRef();
			m_objpool.push_front(pObj);
			bRet = true;
		}

		return bRet;
	}

	HRESULT Hook(IHTMLDocument2* pDoc)
	{
		HRESULT hr = E_FAIL;

		//get the IHTMLWindow5 to register the factory 
		CComPtr<IHTMLWindow2> pWin;
		hr = pDoc->get_parentWindow(&pWin);

		if( SUCCEEDED(hr) )
		{
			CComQIPtr<IHTMLWindow5> pWin5 = pWin;

			if(pWin5 != NULL)
			{
				VARIANT v;
				v.vt = VT_DISPATCH;
				v.pdispVal = this;
				hr = pWin5->put_XMLHttpRequest(v);
			}
		}

		return hr;
	}
};

