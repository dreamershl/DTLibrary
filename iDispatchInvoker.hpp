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

#ifdef _M_IX86

#define BOOST_FT_AUTODETECT_CALLING_CONVENTIONS
#define BOOST_MEM_FN_ENABLE_CDECL
#define BOOST_MEM_FN_ENABLE_STDCALL

#endif

#include <string>
#include <comutil.h>
#include <boost/system/system_error.hpp>

#include "dtcommonfuncs.h"
#include <boost/typeof/typeof.hpp>
#include <boost/fusion/tuple.hpp>

#include "Interpreter.hpp"

using namespace std;

#ifndef DTTRACEMSG_DEBUG
#define DTTRACEMSG_DEBUG  	TRACE

inline std::string DTHRESULT_ERRMSG(HRESULT hr)	
{
	stringstream stream;
	stream << std::hex << std::uppercase << _T("0X") << hr;

	return stream.str();
}

#endif

template<>
struct std::hash<_bstr_t> {
	size_t operator()(const _bstr_t& val) const 
	{
		std::hash<std::wstring> convertor;

		return convertor((wchar_t*)val);
	}
};

namespace DT
{
	#define VALID_PARAM_POITNER(pParam)			if(pParam == NULL) return E_POINTER;

	#define BEGIN_INVOKER(childClass)			public: typedef childClass type; childClass(){   

	#define REG_METHOD_ID(id,strName,func)		m_invoker.register_function(id,strName,&type::func,this);

	#define REG_METHOD_NAME(strName,func)		REG_METHOD_ID( (DISPID)m_invoker.GetInvokerID(strName,false),strName,func);
	#define REG_METHOD_DEFAULT(name,func)		m_defMethodID = (DISPID)REG_METHOD_NAME(name,func)
	#define REG_METHOD(func)					REG_METHOD_NAME(#func,func)

	#define REG_ATTR_BASE(strName)				getAttrName((DISPID)m_invoker.GetInvokerID(strName,false)) = L##strName;
	#define REG_R_ATTR_NAME(strName,attr)		REG_ATTR_BASE(strName) getAttrReaderID((DISPID)m_invoker.GetInvokerID(strName,false))  = REG_METHOD_NAME(strName,get_##attr)
	#define REG_W_ATTR_NAME(strName,attr)		REG_ATTR_BASE(strName) getAttrWriterID((DISPID)m_invoker.GetInvokerID(strName,false))  = REG_METHOD_NAME(strName "_W", put_##attr)

	#define REG_ATTR_NAME(strName,attr)			REG_R_ATTR_NAME(strName,attr) REG_W_ATTR_NAME(strName,attr)

	#define REG_ATTR_NAME_DEFAULT(strName,attr)	m_defMethodID = (DISPID)m_invoker.GetInvokerID(strName,false); REG_ATTR_NAME(strName,attr)

	#define REG_R_ATTR(attr)					REG_R_ATTR_NAME(#attr,attr)
	#define REG_W_ATTR(attr)					REG_W_ATTR_NAME(#attr,attr)
	#define REG_ATTR(attr)						REG_ATTR_NAME(#attr,attr) 
	#define REG_ATTR_DEFAULT(attr)				REG_ATTR_NAME_DEFAULT(#attr,attr) 

	#define REG_EVENT(eName)					REG_W_ATTR(eName) 

	#define END_INVOKER							}

	typedef interpreter_param_parser<std::reverse_iterator<VARIANTARG*>> IDispatchParamTokenizer;
	typedef interpreter<IDispatchParamTokenizer,HRESULT,std::hash<std::string>> IDispatchInterpreter;

	template<>
	template<>
	inline static IDispatchParamTokenizer IDispatchInterpreter::make_param_parser<DISPPARAMS ,IDispatchParamTokenizer>( DISPPARAMS const & params)
	{
		typedef std::reverse_iterator<VARIANTARG*> iter;
		return IDispatchParamTokenizer(iter(params.rgvarg+params.cArgs),iter(params.rgvarg));
	}
	
	template<> 
	template<>
	struct IDispatchParamTokenizer::typecastor<HWND,VARIANTARG>
	{
		inline static HWND cast(VARIANTARG & param)
		{
			HWND retVal = NULL;

			if( V_VT(&param) == VT_EMPTY)
			{
				V_VT(&param) = VT_VOID;
				param.byref = NULL;
			}

			retVal = (HWND)param.byref;

			return retVal;
		}
	};

	template<>
	template<>
	struct IDispatchParamTokenizer::typecastor<bool,VARIANTARG>
	{
		inline static bool cast(VARIANTARG & param)
		{
			if(V_VT(&param) == VT_BOOL)
				return (param.boolVal == VARIANT_TRUE);
			else
				return false; 
		}
	};

	template<> 
	template<>
	struct IDispatchParamTokenizer::typecastor<BSTR,VARIANTARG>
	{
		inline static BSTR cast(VARIANTARG & param)
		{
			if(V_VT(&param) == VT_BSTR)
				return param.bstrVal;
			else
				return NULL;
		}
	};

	template<> 
	template<typename Target>
	struct IDispatchParamTokenizer::typecastor<Target,VARIANTARG>
	{
		inline static Target cast(VARIANTARG & param)
		{
			Target retVal;
			VARTYPE vt = V_VT(&param);

			if( vt != VT_EMPTY && vt != VT_NULL)
			{
				_variant_t var(param);
				retVal = (Target)var;
			}
			else
				retVal = Target();

			return retVal;
		}
	};

	/**
	* out parameter, need the pointer. So need set the type first
	*/
	template<> 
	template<typename Target>
	struct IDispatchParamTokenizer::typecastor<Target*,VARIANTARG>
	{
		inline static Target* cast(VARIANTARG & param)
		{
			Target* retVal = NULL;

			VARTYPE vt = V_VT(&param);

			if( vt == VT_EMPTY || vt == VT_NULL)
			{
				Target defaultVal = Target();

				//use the _variant_t to get the correct VT type
				_variant_t var(defaultVal);
				param = var.Detach();
			}

			retVal = (Target*)(&(param.byref));			
			
			return retVal;
		}
	};

	template<typename iT,REFIID iTIID = __uuidof(iT)>
	class iDispatchInvoker : public iT
	{
	public:
		typedef iT	base_type;
		typedef iDispatchInvoker<iT,iTIID> type;

		iDispatchInvoker()
			:m_dwRef(1)
			,m_IDispatchExHelper(this)
			,m_defMethodID(0)
		{
		};

		virtual ~iDispatchInvoker()
		{
		};

		virtual void CleanUp()
		{
			std::for_each(m_attributes.begin(),m_attributes.end(),[&](ATTRIBUTE_MAP::value_type& item){
				getAttrVal(item.first).Clear();
			});
		}

	public:
		/**
		normally the interface pointer will be shared between the server & client side,
		this is to comply with the COM requirement. 
		*/
		inline void* operator new(size_t size) 
		{
			return ::CoTaskMemAlloc(size);
		}

		inline void* operator new(size_t size, void* pBuf) throw()
		{
			if(!pBuf)
				pBuf = ::CoTaskMemAlloc(size);
			
			return pBuf;
		}

		inline void* operator new[](size_t size) 
		{
			return ::CoTaskMemAlloc(size);
		} 

		inline void operator delete(void* pMemory) 
		{
			::CoTaskMemFree(pMemory);
		}

		inline void operator delete[](void* pMemory) 
		{
			::CoTaskMemFree(pMemory);
		}   

	protected:
		enum {attr_name=0,attr_val,attr_reader,attr_writer};
		typedef tuple<std::wstring,_variant_t,size_t, size_t> ATTRIBUTE;

		typedef unordered_map<size_t,ATTRIBUTE> ATTRIBUTE_MAP;
		ATTRIBUTE_MAP			m_attributes;

		inline _variant_t& getAttrVal(LPCTSTR szAttrName)
		{
			return getAttrVal(m_invoker.GetInvokerID(szAttrName,false));
		}
		
		inline _variant_t& getAttrVal(size_t id)
		{
			return get<attr_val>(m_attributes[id]);
		}

		inline std::wstring& getAttrName(size_t id)
		{
			return get<attr_name>(m_attributes[id]);
		}

		inline size_t& getAttrReaderID(size_t id)
		{
			return  get<attr_reader>(m_attributes[id]);
		}

		inline size_t& getAttrWriterID(size_t id)
		{
			return get<attr_writer>(m_attributes[id]);
		}

		IDispatchInterpreter m_invoker;
		
		/** the name for the default property or method */
		DISPID m_defMethodID;

		long m_dwRef;

	public:
		friend struct IDispatchExHelper;

		/**
		* IDISPATCHEX Interface
		*/
		struct IDispatchExHelper : public IDispatchEx
		{
			virtual HRESULT STDMETHODCALLTYPE GetDispID( 
				/* [in] */  BSTR bstrName,
				/* [in] */ DWORD grfdex,
				/* [out] */ DISPID *pid)
			{
				return pOwner->GetDispID(bstrName,grfdex,pid);
			}

			virtual /* [local] */ HRESULT STDMETHODCALLTYPE InvokeEx( 
				/* [annotation][in] */ 
				DISPID id,
				/* [annotation][in] */ 
				LCID lcid,
				/* [annotation][in] */ 
				WORD wFlags,
				/* [annotation][in] */ 
				DISPPARAMS *pdp,
				/* [annotation][out] */ 
				VARIANT *pvarRes,
				/* [annotation][out] */ 
				EXCEPINFO *pei,
				/* [annotation][unique][in] */ 
				IServiceProvider *pspCaller)
			{
				return pOwner->InvokeEx(id,lcid,wFlags,pdp,pvarRes,pei,pspCaller);
			}

			virtual HRESULT STDMETHODCALLTYPE DeleteMemberByName( 
				/* [in] */ BSTR bstrName,
				/* [in] */ DWORD grfdex)
			{
				return pOwner->DeleteMemberByName(bstrName,grfdex);
			}

			virtual HRESULT STDMETHODCALLTYPE DeleteMemberByDispID( 
				/* [in] */ DISPID id)
			{
				return pOwner->DeleteMemberByDispID(id);
			}

			virtual HRESULT STDMETHODCALLTYPE GetMemberProperties( 
				/* [in] */ DISPID id,
				/* [in] */ DWORD grfdexFetch,
				/* [out] */ __RPC__out DWORD *pgrfdex)
			{
				return pOwner->GetMemberProperties(id,grfdexFetch,pgrfdex);
			}

			virtual HRESULT STDMETHODCALLTYPE GetMemberName( 
				/* [in] */ DISPID id,
				/* [out] */BSTR *pbstrName)
			{
				return pOwner->GetMemberName(id,pbstrName);
			}

			virtual HRESULT STDMETHODCALLTYPE GetNextDispID( 
				/* [in] */ DWORD grfdex,
				/* [in] */ DISPID id,
				/* [out] */DISPID *pid)
			{
				return pOwner->GetNextDispID(grfdex,id,pid);
			}

			virtual HRESULT STDMETHODCALLTYPE GetNameSpaceParent( 
				/* [out] */ __RPC__deref_out_opt IUnknown **ppunk)
			{
				return pOwner->GetNameSpaceParent(ppunk);
			}

			//dummy IDISPATCH interface
		public:
			virtual HRESULT STDMETHODCALLTYPE GetTypeInfoCount( 
				/* [out] */ __RPC__out UINT *pctinfo)
			{
				return pOwner->GetTypeInfoCount(pctinfo);
			}

			virtual HRESULT STDMETHODCALLTYPE GetTypeInfo( 
				/* [in] */ UINT iTInfo,
				/* [in] */ LCID lcid,
				/* [out] */ __RPC__deref_out_opt ITypeInfo **ppTInfo)
			{
				return pOwner->GetTypeInfo(iTInfo,lcid,ppTInfo);
			}

			virtual HRESULT STDMETHODCALLTYPE GetIDsOfNames( 
				/* [in] */ __RPC__in REFIID riid,
				/* [size_is][in] */ __RPC__in_ecount_full(cNames) LPOLESTR *rgszNames,
				/* [range][in] */ __RPC__in_range(0,16384) UINT cNames,
				/* [in] */ LCID lcid,
				/* [size_is][out] */ __RPC__out_ecount_full(cNames) DISPID *rgDispId)
			{
				return pOwner->GetIDsOfNames(riid,rgszNames,cNames,lcid,rgDispId);
			}

			virtual /* [local] */ HRESULT STDMETHODCALLTYPE Invoke( 
				/* [in] */ DISPID dispIdMember,
				/* [in] */ REFIID riid,
				/* [in] */ LCID lcid,
				/* [in] */ WORD wFlags,
				/* [out][in] */ DISPPARAMS *pDispParams,
				/* [out] */ VARIANT *pVarResult,
				/* [out] */ EXCEPINFO *pExcepInfo,
				/* [out] */ UINT *puArgErr)
			{
				return pOwner->Invoke(dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr);
			}


			virtual HRESULT STDMETHODCALLTYPE QueryInterface( 
				/* [in] */ REFIID riid,
				/* [iid_is][out] */ __RPC__deref_out void __RPC_FAR *__RPC_FAR *ppvObject)
			{
				return pOwner->QueryInterface(riid,ppvObject);
			}

			virtual ULONG STDMETHODCALLTYPE AddRef( void)
			{
				return pOwner->AddRef();
			}

			virtual ULONG STDMETHODCALLTYPE Release( void)
			{
				return pOwner->Release();
			}

			IDispatchExHelper(type * owner) : pOwner(owner)
			{
			}


			type* const pOwner;

		};

		IDispatchExHelper m_IDispatchExHelper;

		/**
		* IDISPATCHEx Interface
		*/
		virtual HRESULT STDMETHODCALLTYPE GetDispID( 
			/* [in] */ __RPC__in BSTR bstrName,
			/* [in] */ DWORD grfdex,
			/* [out] */ __RPC__out DISPID *pid)
		{
			return GetIDsOfNames(IID_NULL,&bstrName,1,LOCALE_USER_DEFAULT,pid);
		}
		
		virtual /* [local] */ HRESULT STDMETHODCALLTYPE InvokeEx( 
			/* [annotation][in] */ 
			__in  DISPID id,
			/* [annotation][in] */ 
			__in  LCID lcid,
			/* [annotation][in] */ 
			__in  WORD wFlags,
			/* [annotation][in] */ 
			__in  DISPPARAMS *pdp,
			/* [annotation][out] */ 
			__out_opt  VARIANT *pvarRes,
			/* [annotation][out] */ 
			__out_opt  EXCEPINFO *pei,
			/* [annotation][unique][in] */ 
			__in_opt  IServiceProvider *pspCaller)
		{
			return Invoke(id,IID_NULL,lcid,wFlags,pdp,pvarRes,pei,NULL);
		}
		
		virtual HRESULT STDMETHODCALLTYPE DeleteMemberByName( 
			/* [in] */ __RPC__in BSTR bstrName,
			/* [in] */ DWORD grfdex)
		{
			HRESULT hr = STG_E_UNIMPLEMENTEDFUNCTION;
			
			
			DTTRACEMSG_DEBUG(_T("DeleteMemberByName(%ls,0X%X): Result=%s")
				,bstrName
				,grfdex
				,DTHRESULT_ERRMSG(hr).c_str()
				);

			return hr;
		}
		
		virtual HRESULT STDMETHODCALLTYPE DeleteMemberByDispID( 
			/* [in] */ DISPID id)
		{
			HRESULT hr = STG_E_UNIMPLEMENTEDFUNCTION;
			
			
			DTTRACEMSG_DEBUG(_T("DeleteMemberByDispID(0X%X): Result=%s")
				,id
				,DTHRESULT_ERRMSG(hr).c_str()
				);

			return hr;
		}
		
		virtual HRESULT STDMETHODCALLTYPE GetMemberProperties( 
			/* [in] */ DISPID id,
			/* [in] */ DWORD grfdexFetch,
			/* [out] */ __RPC__out DWORD *pgrfdex)
		{
			HRESULT hr = STG_E_UNIMPLEMENTEDFUNCTION;
			
			
			DTTRACEMSG_DEBUG(_T("GetMemberProperties(0X%X,0X%X,0X%X): Result=%s")
				,id
				,grfdexFetch
				,pgrfdex ? *pgrfdex : 0
				,DTHRESULT_ERRMSG(hr).c_str()
				);

			return hr;
		}
		
		virtual HRESULT STDMETHODCALLTYPE GetMemberName( 
			/* [in] */ DISPID id,
			/* [out] */ __RPC__deref_out_opt BSTR *pbstrName)
		{
			HRESULT hr = STG_E_UNIMPLEMENTEDFUNCTION;
			
			
			DTTRACEMSG_DEBUG(_T("GetMemberName(0X%X,%ls): Result=%s")
				,id
				,pbstrName ? *pbstrName : L""
				,DTHRESULT_ERRMSG(hr).c_str()
				);

			return hr;
		}
		
		virtual HRESULT STDMETHODCALLTYPE GetNextDispID( 
			/* [in] */ DWORD grfdex,
			/* [in] */ DISPID id,
			/* [out] */ __RPC__out DISPID *pid)
		{
			HRESULT hr = STG_E_UNIMPLEMENTEDFUNCTION;
			
			
			DTTRACEMSG_DEBUG(_T("GetNextDispID(0X%X,0X%X,0X%X): Result=%s")
				,grfdex
				,id
				,pid ? *pid : 0
				,DTHRESULT_ERRMSG(hr).c_str()
				);

			return hr;
		}
		
		virtual HRESULT STDMETHODCALLTYPE GetNameSpaceParent( 
			/* [out] */ __RPC__deref_out_opt IUnknown **ppunk)
		{
			HRESULT hr = STG_E_UNIMPLEMENTEDFUNCTION;
			
			
			DTTRACEMSG_DEBUG(_T("GetNameSpaceParent(0X%X): Result=%s")
				,ppunk ? *ppunk : 0
				,DTHRESULT_ERRMSG(hr).c_str()
				);

			return hr;
		}

		/**
		* IDISPATCH Interface
		*/
		// IUnknown
		virtual HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void **ppv)
		{
			*ppv = NULL;
			HRESULT hr = E_NOINTERFACE;

			if (riid == IID_IUnknown || riid == IID_IDispatch) 
				*ppv = static_cast<IDispatch*>(this);
			else if(riid == iTIID)
				*ppv = static_cast<iT*>(this);
			else if(riid == IID_IDispatchEx)
				*ppv = &m_IDispatchExHelper;

			if (*ppv != NULL) 
			{
				AddRef();
				hr = S_OK;
			}

			DTTRACEMSG_DEBUG(_T("QueryInterface(%s): Result=%s")
				,GetGUIDName(riid).c_str()
				,DTHRESULT_ERRMSG(hr).c_str()
				);

			return hr;
		}

		virtual ULONG STDMETHODCALLTYPE AddRef() 
		{ 
			DTTRACEMSG_DEBUG(_T("AddRef(): Ref=%d")
				,m_dwRef+1
				);

			return m_dwRef++;
		};

		virtual ULONG STDMETHODCALLTYPE Release() 
		{ 
			ULONG uRet = --m_dwRef; 
			
			DTTRACEMSG_DEBUG(_T("Release(): Ref=%d")
				,uRet
				);

			if(m_dwRef == 0) 
				delete this; 
			
			return uRet;
		};

		//IDispatch
		virtual HRESULT STDMETHODCALLTYPE GetTypeInfoCount(UINT *pctinfo)
		{
			*pctinfo = 1;

			return S_OK;
		}

		virtual HRESULT STDMETHODCALLTYPE GetTypeInfo(UINT iTInfo, LCID lcid,
			ITypeInfo **ppTInfo)
		{
			return E_FAIL;
		}

		virtual HRESULT STDMETHODCALLTYPE GetIDsOfNames(REFIID riid,
			LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId)
		{
			HRESULT hr = S_OK;
			int nID = 0;

			for (UINT i = 0; i < cNames; i++) 
			{
				_bstr_t strName = rgszNames[i];
				LPCSTR pName = (LPCSTR)strName;
				DISPID key = (DISPID)m_invoker.GetInvokerID(pName,false);

				if(m_invoker.IsRegisteredID(key) ||
					(m_attributes.count(key) > 0)
					)
					rgDispId[i] = key;
				else
				{
					rgDispId[i] = DISPID_UNKNOWN;
					hr = DISP_E_UNKNOWNNAME;
				}

				DTTRACEMSG_DEBUG(_T("query ID of name[%s],Result:ID=0X%X,%s")
				,pName
				,rgDispId[i] 
				,DTHRESULT_ERRMSG(hr).c_str());
			}

			return hr;
		}

		void RaiseException(error_code& err, LPCTSTR szMsg=NULL)
		{
			if(szMsg == NULL)
				szMsg = "";

			throw system_error(err,szMsg);
		}

		/**
		provide the detail description to the caller. refer to the EXCEPINFO MSDN
		for the detail. The child class should override this function in order to 
		provide the more concrete description for the error
		*/
		virtual void ExceptionDetail(EXCEPINFO& info, std::exception& e, int nErrCode = ERROR_UNHANDLED_EXCEPTION )
		{
			_bstr_t bstr;

			//The error code. Error codes should be greater than 1000. 
			//Either this field or the scode field must be filled in; the other must be set to 0.
			info.wCode = 0;
			info.scode = MAKE_HRESULT(SEVERITY_ERROR, FACILITY_ITF, nErrCode);

			info.wReserved = 0;
			info.pvReserved = NULL;
			info.bstrHelpFile = NULL;
			info.dwHelpContext = 0;

			bstr = typeid(this).name();
			info.bstrSource = bstr.Detach();

			bstr = e.what();
			info.bstrDescription = bstr.Detach();
			info.pfnDeferredFillIn = NULL;
		}


		virtual HRESULT STDMETHODCALLTYPE Invoke(DISPID dispIdMember, REFIID riid,
			LCID lcid, WORD wFlags, DISPPARAMS *Params, VARIANT *pVarResult,
			EXCEPINFO *pExcepInfo, UINT *puArgErr)
		{
			/*invoke has 4 types methods
			DISPATCH_METHOD         
			DISPATCH_PROPERTYGET    
			DISPATCH_PROPERTYPUT    
			DISPATCH_PROPERTYPUTREF 

			DISPATCH_CONSTRUCT for the idispatchex
			*/
			HRESULT hr = DISP_E_MEMBERNOTFOUND;
			
			int type = wFlags;
			bool bRetSelf = false;

			DISPPARAMS varResult;
			varResult.cArgs = pVarResult != NULL ? 1 : 0;
			varResult.rgvarg = pVarResult;
					
			if((dispIdMember == DISPID_VALUE && Params->cArgs == 0 && (wFlags & DISPATCH_METHOD) == DISPATCH_METHOD ))
				type = DISPATCH_CONSTRUCT;

			if((type & DISPATCH_CONSTRUCT) == DISPATCH_CONSTRUCT)
			{
				//default method
				if(m_defMethodID != 0)
				{
					Params = &varResult;
					dispIdMember = m_defMethodID;
				}
				else
				{
					//return itself 
					hr = S_OK;
					bRetSelf = true;
					
					pVarResult->vt = VT_DISPATCH;
					pVarResult->pdispVal = this;
				}
			}
			else if((type & DISPATCH_PROPERTYGET) == DISPATCH_PROPERTYGET)
			{
				if(m_attributes.count(dispIdMember) > 0)
				{
					size_t readerID = getAttrReaderID(dispIdMember);
					if(readerID == 0)
					{
						hr = S_OK;
						*pVarResult = getAttrVal(dispIdMember);
					}
					else
					{
						Params = &varResult;
						dispIdMember = (DISPID)readerID;
					}
				}
			}
			else if( ((type & DISPATCH_PROPERTYPUT) == DISPATCH_PROPERTYPUT) || 
					((type & DISPATCH_PROPERTYPUTREF) == DISPATCH_PROPERTYPUTREF)
					)
			{
				if(m_attributes.count(dispIdMember) > 0)
				{
					size_t writerID = getAttrWriterID(dispIdMember);
					if(writerID == 0)
					{
						hr = S_OK;
						getAttrVal(dispIdMember) = Params->rgvarg[0];
					}
					else
						dispIdMember = (DISPID)writerID;
				}
			}
			
			if( hr == DISP_E_MEMBERNOTFOUND && m_invoker.IsRegisteredID(dispIdMember))
			{
				try
				{
					//for some methods, just a getter without any parameters
					if(Params->cArgs == 0)
						Params = &varResult;

					hr = m_invoker.parse_input(dispIdMember,*Params);
				}
				catch(system_error& e)
				{
					if(pExcepInfo != NULL)
						ExceptionDetail(*pExcepInfo,e, e.code().value());

					hr = DISP_E_EXCEPTION;
				}
				catch(std::exception& e)
				{
					if(pExcepInfo != NULL)
						ExceptionDetail(*pExcepInfo,e);

					hr = DISP_E_EXCEPTION;
				}
			}

			DTTRACEMSG_DEBUG(_T("invoke function %s (ID=0X%X) with flag[0X%X]:Result=%s")
				,(bRetSelf ? _T("DISPATCH_CONSTRUCT") : m_invoker.GetInvokerName(dispIdMember).c_str())
				,dispIdMember
				,wFlags
				,DTHRESULT_ERRMSG(hr).c_str()
				);

			return hr;
		}
	};
}

