//https://forums.codeguru.com/showthread.php?232044-IDispatch-Invoke-Send-arguments-by-ref

#include <iostream>
#include <vector>

//COM
#include "comdef.h"
#include "oleacc.h"
#include <atlsafe.h> //ATL

HRESULT AutoWrap(int autoType, VARIANT* pvResult, IDispatch* pDisp, const wchar_t* ptName, int cArgs...) {
    va_list marker;// Begin variable-argument list...
    va_start(marker, cArgs);
    HRESULT hr = S_FALSE;
    if (!pDisp) {
        va_end(marker);
        return hr;
    }
    DISPPARAMS dp = { NULL, NULL, 0, 0 };// Variables used...
    DISPID dispidNamed = DISPID_PROPERTYPUT;
    DISPID dispID;
    hr = pDisp->GetIDsOfNames(IID_NULL, (LPOLESTR*)&ptName, 1, LOCALE_NEUTRAL, &dispID);// Get DISPID for name passed...
    if (FAILED(hr)) {
        va_end(marker);
        return hr;
    }
    VARIANT* pArgs = new VARIANT[cArgs + 1];// Allocate memory for arguments...
    for (int i = 0; i < cArgs; i++) {// Extract arguments...
        pArgs[i] = va_arg(marker, VARIANT);
    }
    dp.cArgs = cArgs;// Build DISPPARAMS
    dp.rgvarg = pArgs;
    if (autoType & DISPATCH_PROPERTYPUT) {// Handle special-case for property-puts!
        dp.cNamedArgs = 1;
        dp.rgdispidNamedArgs = &dispidNamed;
    }
    hr = pDisp->Invoke(dispID, IID_NULL, LOCALE_NEUTRAL, autoType, &dp, pvResult, NULL, NULL);// Make the call!
    if (FAILED(hr)) {
        va_end(marker);
        delete[] pArgs;
        return hr;
    }
    va_end(marker);
    delete[] pArgs;
    return hr;
}

HRESULT RunComNotATLvBSTR()
{
    IDispatchPtr pBedvitComVBADisp = NULL;
    HRESULT hr = 0;
    GUID guid;
    hr = IIDFromString(L"{7a65494f-2a91-415c-9ff6-38f6611675ce}", &guid);//IVBA
    if (FAILED(hr)) {
        return hr;
    }
    hr = CoCreateInstance(guid, NULL, CLSCTX_INPROC_SERVER, IID_IDispatch, (void**)&pBedvitComVBADisp);
    if (FAILED(hr)) {
        return hr;
    }

    //переменные, создаем массив
    _variant_t result, array, disp(pBedvitComVBADisp.GetInterfacePtr());
    SAFEARRAYBOUND bound[2] = { { 3, 1 }, { 2, 1 } }; //3строки, 2 столбца
    array.vt = VT_ARRAY | VT_BSTR; // array of Variants
    array.parray = SafeArrayCreate(VT_BSTR, 2, bound);
    if (!array.parray) {
        return E_POINTER;
    }

    //заполним массив
    BSTR* arr = NULL;
    hr = SafeArrayAccessData(array.parray, (void HUGEP**) & arr);//открываем массив
    if (FAILED(hr)) {
        return hr;
    }
    arr[0] = _bstr_t(L"C").Detach();
    arr[1] = _bstr_t(L"B").Detach();
    arr[2] = _bstr_t(L"A").Detach();

    hr = SafeArrayUnaccessData(array.parray); //закрываем массив
    if (FAILED(hr)) {
        return hr;
    }

    //оборачиваем массив в VT_BYREF, что бы получить результат из метода [in, out]
    VARIANT pvarArray;
    pvarArray.vt = VT_VARIANT | VT_BYREF;
    pvarArray.pvarVal = &array;

    //сортируем//до сортировки 1й столбец (C, B, A)//после сортировки 1й столбец (A, B, C) 
    hr = AutoWrap(DISPATCH_METHOD, &result, disp.pdispVal, L"ArraySortS", 1, pvarArray);
    if (FAILED(hr)) {
        return hr;
    }
    return 0;
}

HRESULT RunComNotATL()
{
    IDispatchPtr pBedvitComVBADisp = NULL;
    HRESULT hr = 0;
    GUID guid;
    hr = IIDFromString(L"{7a65494f-2a91-415c-9ff6-38f6611675ce}", &guid);//IVBA
    if (FAILED(hr)) {
        return hr;
    }
    hr = CoCreateInstance(guid, NULL, CLSCTX_INPROC_SERVER, IID_IDispatch, (void**)&pBedvitComVBADisp);
    if (FAILED(hr)) {
        return hr;
    }

    //переменные, создаем массив
    _variant_t result, array, disp(pBedvitComVBADisp);
	SAFEARRAYBOUND bound[2] = { { 3, 1 }, { 2, 1 } }; //3строки, 2 столбца
    array.vt = VT_ARRAY | VT_VARIANT; // array of Variants
    array.parray = SafeArrayCreate(VT_VARIANT, 2, bound);
	if (!array.parray) {
		return E_POINTER;
	}

    //заполним массив
    VARIANT* arr = NULL;
    hr = SafeArrayAccessData(array.parray, (void HUGEP**) & arr);//открываем массив
    if (FAILED(hr)) {
        return hr;
    }
    arr[0] = _variant_t(L"C").Detach();
    arr[1] = _variant_t(L"B").Detach();
    arr[2] = _variant_t(L"A").Detach();
    hr = SafeArrayUnaccessData(array.parray); //закрываем массив
    if (FAILED(hr)) {
        return hr;
    }

    //оборачиваем массив в VT_BYREF, что бы получить результат из метода [in, out]
    VARIANT pvarArray;
    pvarArray.vt = VT_VARIANT | VT_BYREF;
    pvarArray.pvarVal = &array;

    //сортируем//до сортировки 1й столбец (C, B, A)//после сортировки 1й столбец (A, B, C) 
    hr = AutoWrap(DISPATCH_METHOD, &result, disp.pdispVal, L"ArraySortV", 1, pvarArray);
    if (FAILED(hr)) {
        return hr;
    }
    return 0;
}

HRESULT RunCom()
{
    CComPtr <IDispatch> pBedvitComVBADisp = NULL;
    HRESULT hr = 0;
    GUID guid;
    hr = IIDFromString(L"{7a65494f-2a91-415c-9ff6-38f6611675ce}", &guid);//IVBA
    if (FAILED(hr)) {
        return hr;
    }
    hr = CoCreateInstance(guid, NULL, CLSCTX_INPROC_SERVER, IID_IDispatch, (void**)&pBedvitComVBADisp);
    if (FAILED(hr)) {
        return hr;
    }

    //переменные, создаем массив
    ATL::CComVariant result, disp(pBedvitComVBADisp), array;
    std::vector < ATL::CComSafeArrayBound > bound({ { 3, 1 }, { 2, 1 } });//3строки, 2 столбца
    ATL::CComSafeArray<VARIANT> safeArr(bound.data(), bound.size());
    ////заполним массив
    //LONG alIndex[2] = { 1, 1 };
    //safeArr.MultiDimSetAt(alIndex, ATL::CComVariant(L"A"));
    //alIndex[0] = 2;
    //safeArr.MultiDimSetAt(alIndex, ATL::CComVariant(L"B"));
    //alIndex[0] = 3;
    //safeArr.MultiDimSetAt(alIndex, ATL::CComVariant(L"A"));

    array.vt = VT_ARRAY | safeArr.GetType();
    array.parray = safeArr.Detach();

    //заполним массив
    VARIANT* arr = NULL;
    hr = SafeArrayAccessData(array.parray, (void HUGEP**) & arr);//открываем массив
    if (FAILED(hr)) {
        return hr;
    }
    ATL::CComVariant(L"C").Detach(&arr[0]);
    ATL::CComVariant(L"B").Detach(&arr[1]);
    ATL::CComVariant(L"A").Detach(&arr[2]);
    hr = SafeArrayUnaccessData(array.parray); //закрываем массив
    if (FAILED(hr)) {
        return hr;
    }

    //оборачиваем массив в VT_BYREF, что бы получить результат из метода [in, out]
    VARIANT pvarArray;
    pvarArray.vt = VT_VARIANT | VT_BYREF;
    pvarArray.pvarVal = &array;

    //сортируем//до сортировки 1й столбец (C, B, A)//после сортировки 1й столбец (A, B, C) 
    hr = AutoWrap(DISPATCH_METHOD, &result, disp.pdispVal, L"ArraySortV", 1, pvarArray);
    if (FAILED(hr)) {
        return hr;
    }
    return 0;
}

int main()
{
    HRESULT hr = S_OK;
    hr = OleInitialize(NULL);
    if (FAILED(hr)) {
        MessageBoxW(NULL, _com_error(hr).ErrorMessage(), L"Error", MB_ICONERROR | MB_TOPMOST);
        return hr;
    }

    //вариант без ATL массив VARIANT
    hr = RunComNotATLvBSTR();
    if (FAILED(hr)) {
        MessageBoxW(NULL, _com_error(hr).ErrorMessage(), L"Error", MB_ICONERROR | MB_TOPMOST);
        OleUninitialize();
        return hr;
    }

    //вариант без ATL массив BSTR
    hr = RunComNotATL();
    if (FAILED(hr)) {
        MessageBoxW(NULL, _com_error(hr).ErrorMessage(), L"Error", MB_ICONERROR | MB_TOPMOST);
        OleUninitialize();
        return hr;
    }

    //вариант с ATL
    hr = RunCom();
    if (FAILED(hr)) {
        MessageBoxW(NULL, _com_error(hr).ErrorMessage(), L"Error", MB_ICONERROR | MB_TOPMOST);
    }
    OleUninitialize();
    return hr;
}