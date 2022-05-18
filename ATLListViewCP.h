
#pragma once
template <class T>
class CProxy_IListControlEvents
    : public IConnectionPointImpl<T, &__uuidof(_IListControlEvents)> {
    public:
    
		
		HRESULT Fire_KeyPress(SHORT* KeyAscii) {
        CComVariant varResult;
        T* pT = static_cast<T*>(this);
        int nConnectionIndex;
        CComVariant* pvars = new CComVariant[1];
        int nConnections = m_vec.GetSize();

        for (nConnectionIndex = 0; nConnectionIndex < nConnections;
             nConnectionIndex++) {
            pT->Lock();
            CComPtr<IUnknown> sp = m_vec.GetAt(nConnectionIndex);
            pT->Unlock();
            IDispatch* pDispatch = reinterpret_cast<IDispatch*>(sp.p);
            if (pDispatch != NULL) {
                VariantClear(&varResult);
                pvars[0].byref = KeyAscii;
                pvars[0].vt = VT_I2 | VT_BYREF;
                DISPPARAMS disp = {pvars, NULL, 1, 0};
                pDispatch->Invoke(18, IID_NULL, LOCALE_USER_DEFAULT,
                    DISPATCH_METHOD, &disp, &varResult, NULL, NULL);
            }
        }
        delete[] pvars;
        return varResult.scode;
    }

    HRESULT Fire_KeyUp(SHORT Key, SHORT Shift) {
        CComVariant varResult;
        T* pT = static_cast<T*>(this);
        int nConnectionIndex;
        CComVariant* pvars = new CComVariant[2];
        int nConnections = m_vec.GetSize();

        for (nConnectionIndex = 0; nConnectionIndex < nConnections;
             nConnectionIndex++) {
            pT->Lock();
            CComPtr<IUnknown> sp = m_vec.GetAt(nConnectionIndex);
            pT->Unlock();
            IDispatch* pDispatch = reinterpret_cast<IDispatch*>(sp.p);
            if (pDispatch != NULL) {
                VariantClear(&varResult);
                pvars[1] = Key;
                pvars[0] = Shift;
                DISPPARAMS disp = {pvars, NULL, 2, 0};
                pDispatch->Invoke(17, IID_NULL, LOCALE_USER_DEFAULT,
                    DISPATCH_METHOD, &disp, &varResult, NULL, NULL);
            }
        }
        delete[] pvars;
        return varResult.scode;
    }

    HRESULT Fire_KeyDown(SHORT Key, SHORT Shift) {
        CComVariant varResult;
        T* pT = static_cast<T*>(this);
        int nConnectionIndex;
        CComVariant* pvars = new CComVariant[2];
        int nConnections = m_vec.GetSize();

        for (nConnectionIndex = 0; nConnectionIndex < nConnections;
             nConnectionIndex++) {
            pT->Lock();
            CComPtr<IUnknown> sp = m_vec.GetAt(nConnectionIndex);
            pT->Unlock();
            IDispatch* pDispatch = reinterpret_cast<IDispatch*>(sp.p);
            if (pDispatch != NULL) {
                VariantClear(&varResult);
                pvars[1] = Key;
                pvars[0] = Shift;
                DISPPARAMS disp = {pvars, NULL, 2, 0};
                pDispatch->Invoke(16, IID_NULL, LOCALE_USER_DEFAULT,
                    DISPATCH_METHOD, &disp, &varResult, NULL, NULL);
            }
        }
        delete[] pvars;
        return varResult.scode;
    }

    HRESULT Fire_DblClick() {
        CComVariant varResult;
        T* pT = static_cast<T*>(this);
        int nConnectionIndex;
        int nConnections = m_vec.GetSize();

        for (nConnectionIndex = 0; nConnectionIndex < nConnections;
             nConnectionIndex++) {
            pT->Lock();
            CComPtr<IUnknown> sp = m_vec.GetAt(nConnectionIndex);
            pT->Unlock();
            IDispatch* pDispatch = reinterpret_cast<IDispatch*>(sp.p);
            if (pDispatch != NULL) {
                VariantClear(&varResult);
                DISPPARAMS disp = {NULL, NULL, 0, 0};
                pDispatch->Invoke(15, IID_NULL, LOCALE_USER_DEFAULT,
                    DISPATCH_METHOD, &disp, &varResult, NULL, NULL);
            }
        }
        return varResult.scode;
    }
    HRESULT Fire_ColumnClick(IColumnHeader* whichHeader) {
        HRESULT hr = S_OK;
        T* pThis = static_cast<T*>(this);
        int cConnections = m_vec.GetSize();

        for (int iConnection = 0; iConnection < cConnections; iConnection++) {
            pThis->Lock();
            CComPtr<IUnknown> punkConnection = m_vec.GetAt(iConnection);
            pThis->Unlock();

            IDispatch* pConnection = static_cast<IDispatch*>(punkConnection.p);

            if (pConnection) {
                CComVariant avarParams[1];
                avarParams[0] = whichHeader;
                CComVariant varResult;

                DISPPARAMS params = {avarParams, NULL, 1, 0};
                hr = pConnection->Invoke(1, IID_NULL, LOCALE_USER_DEFAULT,
                    DISPATCH_METHOD, &params, &varResult, NULL, NULL);
            }
        }
        return hr;
    }
    HRESULT Fire_ColumnRightClick(VARIANT_BOOL* DoDefault, SHORT Shift, FLOAT x,
        FLOAT y, IColumnHeader* whichHeader) {
        HRESULT hr = S_OK;
        T* pThis = static_cast<T*>(this);
        int cConnections = m_vec.GetSize();

        for (int iConnection = 0; iConnection < cConnections; iConnection++) {
            pThis->Lock();
            CComPtr<IUnknown> punkConnection = m_vec.GetAt(iConnection);
            pThis->Unlock();

            IDispatch* pConnection = static_cast<IDispatch*>(punkConnection.p);

            if (pConnection) {
                CComVariant avarParams[5];
                avarParams[4].byref = DoDefault;
                avarParams[4].vt = VT_BOOL | VT_BYREF;
                avarParams[3] = Shift;
                avarParams[3].vt = VT_I2;
                avarParams[2] = x;
                avarParams[2].vt = VT_R4;
                avarParams[1] = y;
                avarParams[1].vt = VT_R4;
                avarParams[0] = whichHeader;
                CComVariant varResult;

                DISPPARAMS params = {avarParams, NULL, 5, 0};
                hr = pConnection->Invoke(2, IID_NULL, LOCALE_USER_DEFAULT,
                    DISPATCH_METHOD, &params, &varResult, NULL, NULL);
            }
        }
        return hr;
    }
    HRESULT Fire_ColumnAdded(ColumnHeader* colAdded) {
        HRESULT hr = S_OK;
        T* pThis = static_cast<T*>(this);
        int cConnections = m_vec.GetSize();

        for (int iConnection = 0; iConnection < cConnections; iConnection++) {
            pThis->Lock();
            CComPtr<IUnknown> punkConnection = m_vec.GetAt(iConnection);
            pThis->Unlock();

            IDispatch* pConnection = static_cast<IDispatch*>(punkConnection.p);

            if (pConnection) {
                CComVariant avarParams[1];
                avarParams[0] = colAdded;
                CComVariant varResult;

                DISPPARAMS params = {avarParams, NULL, 1, 0};
                hr = pConnection->Invoke(3, IID_NULL, LOCALE_USER_DEFAULT,
                    DISPATCH_METHOD, &params, &varResult, NULL, NULL);
            }
        }
        return hr;
    }
    HRESULT Fire_ColumnRemoved(ColumnHeader* colRemoved) {
        HRESULT hr = S_OK;
        T* pThis = static_cast<T*>(this);
        int cConnections = m_vec.GetSize();

        for (int iConnection = 0; iConnection < cConnections; iConnection++) {
            pThis->Lock();
            CComPtr<IUnknown> punkConnection = m_vec.GetAt(iConnection);
            pThis->Unlock();

            IDispatch* pConnection = static_cast<IDispatch*>(punkConnection.p);

            if (pConnection) {
                CComVariant avarParams[1];
                avarParams[0] = colRemoved;
                CComVariant varResult;

                DISPPARAMS params = {avarParams, NULL, 1, 0};
                hr = pConnection->Invoke(4, IID_NULL, LOCALE_USER_DEFAULT,
                    DISPATCH_METHOD, &params, &varResult, NULL, NULL);
            }
        }
        return hr;
    }
    HRESULT Fire_ItemSelectionChanged(IListItem* item, VARIANT_BOOL SelState) {
        HRESULT hr = S_OK;
        T* pThis = static_cast<T*>(this);
        int cConnections = m_vec.GetSize();

        for (int iConnection = 0; iConnection < cConnections; iConnection++) {
            pThis->Lock();
            CComPtr<IUnknown> punkConnection = m_vec.GetAt(iConnection);
            pThis->Unlock();

            IDispatch* pConnection = static_cast<IDispatch*>(punkConnection.p);

            if (pConnection) {
                CComVariant avarParams[2];
                avarParams[1] = item;
                avarParams[0] = SelState;
                avarParams[0].vt = VT_BOOL;
                CComVariant varResult;

                DISPPARAMS params = {avarParams, NULL, 2, 0};
                hr = pConnection->Invoke(5, IID_NULL, LOCALE_USER_DEFAULT,
                    DISPATCH_METHOD, &params, &varResult, NULL, NULL);
            }
        }
        return hr;
    }
    HRESULT Fire_ItemClick(ListItem* Item) {
        HRESULT hr = S_OK;
        T* pThis = static_cast<T*>(this);
        int cConnections = m_vec.GetSize();

        for (int iConnection = 0; iConnection < cConnections; iConnection++) {
            pThis->Lock();
            CComPtr<IUnknown> punkConnection = m_vec.GetAt(iConnection);
            pThis->Unlock();

            IDispatch* pConnection = static_cast<IDispatch*>(punkConnection.p);

            if (pConnection) {
                CComVariant avarParams[1];
                avarParams[0].vt = VT_BYREF | VT_DISPATCH;
                avarParams[0].byref = Item;
                CComVariant varResult;

                DISPPARAMS params = {avarParams, NULL, 1, 0};
                hr = pConnection->Invoke(6, IID_NULL, LOCALE_USER_DEFAULT,
                    DISPATCH_METHOD, &params, &varResult, NULL, NULL);
            }
        }
        return hr;
    }
    HRESULT Fire_ItemClickEx(ListItem* Item, short SubItemIndex) {
        HRESULT hr = S_OK;
        T* pThis = static_cast<T*>(this);
        int cConnections = m_vec.GetSize();

        for (int iConnection = 0; iConnection < cConnections; iConnection++) {
            pThis->Lock();
            CComPtr<IUnknown> punkConnection = m_vec.GetAt(iConnection);
            pThis->Unlock();

            IDispatch* pConnection = static_cast<IDispatch*>(punkConnection.p);

            if (pConnection) {
                CComVariant avarParams[2];
                avarParams[1].vt = VT_BYREF | VT_DISPATCH;
                avarParams[1].byref = Item;
                avarParams[0] = SubItemIndex;
                avarParams[0].vt = VT_I2;
                CComVariant varResult;

                DISPPARAMS params = {avarParams, NULL, 2, 0};
                hr = pConnection->Invoke(7, IID_NULL, LOCALE_USER_DEFAULT,
                    DISPATCH_METHOD, &params, &varResult, NULL, NULL);
            }
        }
        return hr;
    }
    HRESULT Fire_ItemClickRight(ListItem* item, SHORT subItemIndex) {
        HRESULT hr = S_OK;
        T* pThis = static_cast<T*>(this);
        int cConnections = m_vec.GetSize();

        for (int iConnection = 0; iConnection < cConnections; iConnection++) {
            pThis->Lock();
            CComPtr<IUnknown> punkConnection = m_vec.GetAt(iConnection);
            pThis->Unlock();

            IDispatch* pConnection = static_cast<IDispatch*>(punkConnection.p);

            if (pConnection) {
                CComVariant avarParams[2];
                avarParams[1].vt = VT_BYREF | VT_DISPATCH;
                avarParams[1].byref = item;
                avarParams[0] = subItemIndex;
                avarParams[0].vt = VT_I2;
                CComVariant varResult;

                DISPPARAMS params = {avarParams, NULL, 2, 0};
                hr = pConnection->Invoke(8, IID_NULL, LOCALE_USER_DEFAULT,
                    DISPATCH_METHOD, &params, &varResult, NULL, NULL);
            }
        }
        return hr;
    }
    HRESULT Fire_Click() {
        HRESULT hr = S_OK;
        T* pThis = static_cast<T*>(this);
        int cConnections = m_vec.GetSize();

        for (int iConnection = 0; iConnection < cConnections; iConnection++) {
            pThis->Lock();
            CComPtr<IUnknown> punkConnection = m_vec.GetAt(iConnection);
            pThis->Unlock();

            IDispatch* pConnection = static_cast<IDispatch*>(punkConnection.p);

            if (pConnection) {
                CComVariant varResult;

                DISPPARAMS params = {NULL, NULL, 0, 0};
                hr = pConnection->Invoke(9, IID_NULL, LOCALE_USER_DEFAULT,
                    DISPATCH_METHOD, &params, &varResult, NULL, NULL);
            }
        }
        return hr;
    }
    HRESULT Fire_ClickEx(LONG x, LONG y) {
        HRESULT hr = S_OK;
        T* pThis = static_cast<T*>(this);
        int cConnections = m_vec.GetSize();

        for (int iConnection = 0; iConnection < cConnections; iConnection++) {
            pThis->Lock();
            CComPtr<IUnknown> punkConnection = m_vec.GetAt(iConnection);
            pThis->Unlock();

            IDispatch* pConnection = static_cast<IDispatch*>(punkConnection.p);

            if (pConnection) {
                CComVariant avarParams[2];
                avarParams[1] = x;
                avarParams[1].vt = VT_I4;
                avarParams[0] = y;
                avarParams[0].vt = VT_I4;
                CComVariant varResult;

                DISPPARAMS params = {avarParams, NULL, 2, 0};
                hr = pConnection->Invoke(10, IID_NULL, LOCALE_USER_DEFAULT,
                    DISPATCH_METHOD, &params, &varResult, NULL, NULL);
            }
        }
        return hr;
    }
    HRESULT Fire_ItemClicked(ListItem* Item, vbMouseButtonConstants Button) {
        HRESULT hr = S_OK;
        T* pThis = static_cast<T*>(this);
        int cConnections = m_vec.GetSize();

        for (int iConnection = 0; iConnection < cConnections; iConnection++) {
            pThis->Lock();
            CComPtr<IUnknown> punkConnection = m_vec.GetAt(iConnection);
            pThis->Unlock();

            IDispatch* pConnection = static_cast<IDispatch*>(punkConnection.p);

            if (pConnection) {
                CComVariant avarParams[2];
                avarParams[1] = Item;
                avarParams[0] = Button;
                CComVariant varResult;

                DISPPARAMS params = {avarParams, NULL, 2, 0};
                hr = pConnection->Invoke(11, IID_NULL, LOCALE_USER_DEFAULT,
                    DISPATCH_METHOD, &params, &varResult, NULL, NULL);
            }
        }
        return hr;
    }
    HRESULT Fire_MouseMove(SHORT button, SHORT Shift, FLOAT x, FLOAT y) {
        HRESULT hr = S_OK;
        T* pThis = static_cast<T*>(this);
        int cConnections = m_vec.GetSize();

        for (int iConnection = 0; iConnection < cConnections; iConnection++) {
            pThis->Lock();
            CComPtr<IUnknown> punkConnection = m_vec.GetAt(iConnection);
            pThis->Unlock();

            IDispatch* pConnection = static_cast<IDispatch*>(punkConnection.p);

            if (pConnection) {
                CComVariant avarParams[4];
                avarParams[3] = button;
                avarParams[3].vt = VT_I2;
                avarParams[2] = Shift;
                avarParams[2].vt = VT_I2;
                avarParams[1] = x;
                avarParams[1].vt = VT_R4;
                avarParams[0] = y;
                avarParams[0].vt = VT_R4;
                CComVariant varResult;

                DISPPARAMS params = {avarParams, NULL, 4, 0};
                hr = pConnection->Invoke(12, IID_NULL, LOCALE_USER_DEFAULT,
                    DISPATCH_METHOD, &params, &varResult, NULL, NULL);
            }
        }
        return hr;
    }
    HRESULT Fire_MouseDown(SHORT button, SHORT Shift, FLOAT x, FLOAT y) {
        HRESULT hr = S_OK;
        T* pThis = static_cast<T*>(this);
        int cConnections = m_vec.GetSize();

        for (int iConnection = 0; iConnection < cConnections; iConnection++) {
            pThis->Lock();
            CComPtr<IUnknown> punkConnection = m_vec.GetAt(iConnection);
            pThis->Unlock();

            IDispatch* pConnection = static_cast<IDispatch*>(punkConnection.p);

            if (pConnection) {
                CComVariant avarParams[4];
                avarParams[3] = button;
                avarParams[3].vt = VT_I2;
                avarParams[2] = Shift;
                avarParams[2].vt = VT_I2;
                avarParams[1] = x;
                avarParams[1].vt = VT_R4;
                avarParams[0] = y;
                avarParams[0].vt = VT_R4;
                CComVariant varResult;

                DISPPARAMS params = {avarParams, NULL, 4, 0};
                hr = pConnection->Invoke(13, IID_NULL, LOCALE_USER_DEFAULT,
                    DISPATCH_METHOD, &params, &varResult, NULL, NULL);
            }
        }
        return hr;
    }
    HRESULT Fire_MouseUp(SHORT button, SHORT Shift, FLOAT x, FLOAT y) {
        HRESULT hr = S_OK;
        T* pThis = static_cast<T*>(this);
        int cConnections = m_vec.GetSize();

        for (int iConnection = 0; iConnection < cConnections; iConnection++) {
            pThis->Lock();
            CComPtr<IUnknown> punkConnection = m_vec.GetAt(iConnection);
            pThis->Unlock();

            IDispatch* pConnection = static_cast<IDispatch*>(punkConnection.p);

            if (pConnection) {
                CComVariant avarParams[4];
                avarParams[3] = button;
                avarParams[3].vt = VT_I2;
                avarParams[2] = Shift;
                avarParams[2].vt = VT_I2;
                avarParams[1] = x;
                avarParams[1].vt = VT_R4;
                avarParams[0] = y;
                avarParams[0].vt = VT_R4;
                CComVariant varResult;

                DISPPARAMS params = {avarParams, NULL, 4, 0};
                hr = pConnection->Invoke(14, IID_NULL, LOCALE_USER_DEFAULT,
                    DISPATCH_METHOD, &params, &varResult, NULL, NULL);
            }
        }
        return hr;
    }
};

template <class T>
class CProxy_IColumnHeaderEvents
    : public IConnectionPointImpl<T, &DIID__IColumnHeaderEvents,
          CComDynamicUnkArray> {
    // Warning this class may be recreated by the wizard.
    public:
    HRESULT Fire_ColumnClick(IColumnHeader* whichHeader) {
        CComVariant varResult;
        T* pT = static_cast<T*>(this);
        int nConnectionIndex;
        CComVariant* pvars = new CComVariant[1];
        int nConnections = m_vec.GetSize();

        for (nConnectionIndex = 0; nConnectionIndex < nConnections;
             nConnectionIndex++) {
            pT->Lock();
            CComPtr<IUnknown> sp = m_vec.GetAt(nConnectionIndex);
            pT->Unlock();
            IDispatch* pDispatch = reinterpret_cast<IDispatch*>(sp.p);
            if (pDispatch != NULL) {
                VariantClear(&varResult);
                pvars[0] = whichHeader;
                DISPPARAMS disp = {pvars, NULL, 1, 0};
                pDispatch->Invoke(0x1, IID_NULL, LOCALE_USER_DEFAULT,
                    DISPATCH_METHOD, &disp, &varResult, NULL, NULL);
            }
        }
        delete[] pvars;
        return varResult.scode;
    }
    HRESULT Fire_MouseDown(SHORT Button, SHORT Shift, FLOAT x, FLOAT y) {
        CComVariant varResult;
        T* pT = static_cast<T*>(this);
        int nConnectionIndex;
        CComVariant* pvars = new CComVariant[4];
        int nConnections = m_vec.GetSize();

        for (nConnectionIndex = 0; nConnectionIndex < nConnections;
             nConnectionIndex++) {
            pT->Lock();
            CComPtr<IUnknown> sp = m_vec.GetAt(nConnectionIndex);
            pT->Unlock();
            IDispatch* pDispatch = reinterpret_cast<IDispatch*>(sp.p);
            if (pDispatch != NULL) {
                VariantClear(&varResult);
                pvars[3] = Button;
                pvars[2] = Shift;
                pvars[1] = x;
                pvars[0] = y;
                DISPPARAMS disp = {pvars, NULL, 4, 0};
                pDispatch->Invoke(0x2, IID_NULL, LOCALE_USER_DEFAULT,
                    DISPATCH_METHOD, &disp, &varResult, NULL, NULL);
            }
        }
        delete[] pvars;
        return varResult.scode;
    }
    HRESULT Fire_MouseUp(SHORT Button, SHORT Shift, FLOAT x, FLOAT y) {
        CComVariant varResult;
        T* pT = static_cast<T*>(this);
        int nConnectionIndex;
        CComVariant* pvars = new CComVariant[4];
        int nConnections = m_vec.GetSize();

        for (nConnectionIndex = 0; nConnectionIndex < nConnections;
             nConnectionIndex++) {
            pT->Lock();
            CComPtr<IUnknown> sp = m_vec.GetAt(nConnectionIndex);
            pT->Unlock();
            IDispatch* pDispatch = reinterpret_cast<IDispatch*>(sp.p);
            if (pDispatch != NULL) {
                VariantClear(&varResult);
                pvars[3] = Button;
                pvars[2] = Shift;
                pvars[1] = x;
                pvars[0] = y;
                DISPPARAMS disp = {pvars, NULL, 4, 0};
                pDispatch->Invoke(0x3, IID_NULL, LOCALE_USER_DEFAULT,
                    DISPATCH_METHOD, &disp, &varResult, NULL, NULL);
            }
        }
        delete[] pvars;
        return varResult.scode;
    }
    HRESULT Fire_MouseMove(SHORT Button, SHORT Shift, FLOAT x, FLOAT y) {
        CComVariant varResult;
        T* pT = static_cast<T*>(this);
        int nConnectionIndex;
        CComVariant* pvars = new CComVariant[4];
        int nConnections = m_vec.GetSize();

        for (nConnectionIndex = 0; nConnectionIndex < nConnections;
             nConnectionIndex++) {
            pT->Lock();
            CComPtr<IUnknown> sp = m_vec.GetAt(nConnectionIndex);
            pT->Unlock();
            IDispatch* pDispatch = reinterpret_cast<IDispatch*>(sp.p);
            if (pDispatch != NULL) {
                VariantClear(&varResult);
                pvars[3] = Button;
                pvars[2] = Shift;
                pvars[1] = x;
                pvars[0] = y;
                DISPPARAMS disp = {pvars, NULL, 4, 0};
                pDispatch->Invoke(0x4, IID_NULL, LOCALE_USER_DEFAULT,
                    DISPATCH_METHOD, &disp, &varResult, NULL, NULL);
            }
        }
        delete[] pvars;
        return varResult.scode;
    }
    HRESULT Fire_MouseEvent(LONG iMsg, SHORT Button, SHORT Shift, FLOAT x,
        FLOAT y, VARIANT_BOOL* DoDefault) {
        CComVariant varResult;
        T* pT = static_cast<T*>(this);
        int nConnectionIndex;
        CComVariant* pvars = new CComVariant[6];
        int nConnections = m_vec.GetSize();

        for (nConnectionIndex = 0; nConnectionIndex < nConnections;
             nConnectionIndex++) {
            pT->Lock();
            CComPtr<IUnknown> sp = m_vec.GetAt(nConnectionIndex);
            pT->Unlock();
            IDispatch* pDispatch = reinterpret_cast<IDispatch*>(sp.p);
            if (pDispatch != NULL) {
                VariantClear(&varResult);
                pvars[5] = iMsg;
                pvars[4] = Button;
                pvars[3] = Shift;
                pvars[2] = x;
                pvars[1] = y;
                pvars[0] = DoDefault;
                DISPPARAMS disp = {pvars, NULL, 6, 0};
                pDispatch->Invoke(0x5, IID_NULL, LOCALE_USER_DEFAULT,
                    DISPATCH_METHOD, &disp, &varResult, NULL, NULL);
            }
        }
        delete[] pvars;
        return varResult.scode;
    }
};
