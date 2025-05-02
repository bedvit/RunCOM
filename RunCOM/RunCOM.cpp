// ConsoleApplication1.cpp : Этот файл содержит функцию "main". Здесь начинается и заканчивается выполнение программы.
//
//https://forums.codeguru.com/showthread.php?232044-IDispatch-Invoke-Send-arguments-by-ref

#include <iostream>
#include <vector>

//COM
#include "comdef.h"
#include <atlsafe.h>
#include "oleacc.h"


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
    ATL::CComVariant(L"A").Detach(&arr[0]);
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

    //сортируем//до сортировки 1й столбец (A, B, A)//после сортировки 1й столбец (A, A, B) 
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
    hr = RunCom();
    if (FAILED(hr)) {
        MessageBoxW(NULL, _com_error(hr).ErrorMessage(), L"Error", MB_ICONERROR | MB_TOPMOST);
        //return hr;
    }
    OleUninitialize();
    return hr;
}

// Запуск программы: CTRL+F5 или меню "Отладка" > "Запуск без отладки"
// Отладка программы: F5 или меню "Отладка" > "Запустить отладку"

// Советы по началу работы 
//   1. В окне обозревателя решений можно добавлять файлы и управлять ими.
//   2. В окне Team Explorer можно подключиться к системе управления версиями.
//   3. В окне "Выходные данные" можно просматривать выходные данные сборки и другие сообщения.
//   4. В окне "Список ошибок" можно просматривать ошибки.
//   5. Последовательно выберите пункты меню "Проект" > "Добавить новый элемент", чтобы создать файлы кода, или "Проект" > "Добавить существующий элемент", чтобы добавить в проект существующие файлы кода.
//   6. Чтобы снова открыть этот проект позже, выберите пункты меню "Файл" > "Открыть" > "Проект" и выберите SLN-файл.
