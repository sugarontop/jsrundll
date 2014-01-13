#pragma once

class CActiveScriptSite : public IActiveScriptSite
			#ifdef __Script_Debug__
				   , public IActiveScriptSiteDebug
			#endif
{
	private:
		ULONG m_uRef;
		EXCEPINFO ei_;
		DWORD ErrLineNo_;
		LONG ErrLinePos_;
		IActiveScript* ac_;
	public :
		
		EXCEPINFO Error(){ return ei_; }
		UINT ErrLineNo(){ return (UINT)ErrLineNo_; } 
		UINT ErrPos(){ return (UINT)ErrLinePos_; }

	public:
		CActiveScriptSite(IActiveScript* ac):m_uRef(0),ErrLineNo_(0),ErrLinePos_(0),ac_(ac)
			#ifdef __Script_Debug__
				,mpDebugMgr(0)
				,mpDebugApp(0)
				,mpDebugDoc(0)
				,mAppCookie(0)
			#endif
		{
			ZeroMemory( &ei_, sizeof(EXCEPINFO) );
		}

		~CActiveScriptSite()
		{
			if ( ei_.bstrDescription ) ::SysFreeString( ei_.bstrDescription ); 
			if ( ei_.bstrHelpFile ) ::SysFreeString( ei_.bstrHelpFile );
			if ( ei_.bstrSource ) ::SysFreeString( ei_.bstrSource );
		}

		void NewStaticObject( LPOLESTR* name, IDispatch** StaticObject, int cStatinObject )
		{
			for( int i = 0; i < cStatinObject; i++ )
			{
				mapStaticObject_[ name[i] ] = StaticObject[i];
			}

		}
		void ErrorReset()
		{
			if ( ei_.bstrDescription ) ::SysFreeString( ei_.bstrDescription ); 
			if ( ei_.bstrHelpFile ) ::SysFreeString( ei_.bstrHelpFile );
			if ( ei_.bstrSource ) ::SysFreeString( ei_.bstrSource );


			ZeroMemory( &ei_, sizeof(EXCEPINFO) );
		}

		std::map<std::wstring, IDispatch*> mapStaticObject_;

		STDMETHODIMP_(ULONG) AddRef()
		{
			return InterlockedIncrement(&m_uRef);
		};
		STDMETHODIMP_(ULONG) Release()
		{
			ULONG ulVal = InterlockedDecrement(&m_uRef);
			if(ulVal > 0)
				return ulVal;

			delete this;
			return ulVal;
		};
		STDMETHODIMP QueryInterface(REFIID riid,LPVOID *ppvOut)
		{
			if(*ppvOut)
				*ppvOut = NULL;

			if(IsEqualIID(riid,IID_IActiveScriptSite))
				*ppvOut = (IActiveScriptSite*)this;
			else if(IsEqualIID(riid,IID_IUnknown))
				*ppvOut = this;
		#ifdef __Script_Debug__
			else if (IsEqualIID(riid,IID_IActiveScriptSiteDebug))
				*ppvOut = (IActiveScriptSiteDebug*)this;				
		#endif	
			else
				return E_NOINTERFACE;

			AddRef();
			return S_OK;
		};

		STDMETHODIMP OnScriptError(IActiveScriptError *pscripterror)
		{
			HRESULT hRes = S_OK;
			if(!pscripterror)
				return E_POINTER;

     
			hRes = pscripterror->GetExceptionInfo(&ei_);
      	  
			DWORD d1;
			return pscripterror->GetSourcePosition( &d1, &ErrLineNo_, &ErrLinePos_ );
			
		}

		STDMETHODIMP GetItemInfo(LPCOLESTR pstrName,DWORD dwReturnMask,	IUnknown **ppiunkItem,ITypeInfo **ppti);

		STDMETHODIMP GetDocVersionString(BSTR *pbstrVersion){return E_NOTIMPL;}
		
		STDMETHODIMP OnStateChange(SCRIPTSTATE ssScriptState){return S_OK;}
		STDMETHODIMP OnScriptTerminate(const VARIANT *pvarResult,const EXCEPINFO *pexcepinfo){return S_OK;}
		STDMETHODIMP OnEnterScript(void){return S_OK;}
		STDMETHODIMP OnLeaveScript(void){return S_OK;}
	
		//IActiveScriptSite
		STDMETHODIMP GetLCID(LCID *plcid){*plcid = LOCALE_USER_DEFAULT;return S_OK;}

		CComVariant window_;
		DWORD_PTR  pdwSourceContext;
		//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
		#ifdef __Script_Debug__
			// --- IActiveScriptSiteDebug methods ---
			STDMETHODIMP GetDocumentContextFromPosition(/*[in]*/ DWORD dwSourceContext,
				/*[in]*/ ULONG uCharacterOffset, /*[in]*/ ULONG uNumChars,
				/*[out]*/ IDebugDocumentContext** ppsc); 
			STDMETHODIMP GetApplication(/*[out]*/ IDebugApplication **ppda);
			STDMETHODIMP GetRootApplicationNode(/*[out]*/ IDebugApplicationNode** ppdanRoot); 
			STDMETHODIMP OnScriptErrorDebug(/*[in]*/ IActiveScriptErrorDebug *pErrorDebug,
				/*[out]*/ BOOL *pfEnterDebugger,/*[out]*/ BOOL *pfCallOnScriptErrorWhenContinuing);
					/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
			IProcessDebugManager*	mpDebugMgr;
			IDebugApplication*		mpDebugApp;
			IDebugDocumentHelper*	mpDebugDoc;
			DWORD					mAppCookie;
			

			CComPtr<IDebugDocumentHost> dochost;
			HRESULT CreateScriptDebugger();
			void ReleaseScriptDebugger();
			HRESULT CreateDocumentForDebugger(BSTR scripts);
			void ReleaseDebugDocument();
		#endif
};

