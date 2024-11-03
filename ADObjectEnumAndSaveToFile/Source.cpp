#include <windows.h>
#include <activeds.h>
#include <iads.h>
#include <iostream>
#include <iomanip>
#include <fstream>
#include <comdef.h>
#include <sddl.h>
#include <string>
#include <vector>
using namespace std;

#include <objbase.h>

#pragma comment(lib, "activeds.lib")
#pragma comment(lib, "Adsiid.lib")

void SaveObjectAttributesOld(IADs* pADs, std::ofstream& file) {
    VARIANT var;
    HRESULT hr;
    IADsPropertyList* pPropList;

    hr = pADs->QueryInterface(IID_IADsPropertyList, (void**)&pPropList);
    if (SUCCEEDED(hr)) {
        long lCount;
        hr = pPropList->get_PropertyCount(&lCount);
        if (SUCCEEDED(hr)) {
            for (long i = 0; i < lCount; ++i) {
                VARIANT var;
                hr = pPropList->Next(&var);
                if (SUCCEEDED(hr)) {
                    BSTR bstrName;
                    hr = V_DISPATCH(&var)->QueryInterface(IID_IADsPropertyEntry, (void**)&bstrName);
                    if (SUCCEEDED(hr)) {
                        file << _com_util::ConvertBSTRToString(bstrName) << ": ";
                        hr = pADs->Get(bstrName, &var);
                        if (SUCCEEDED(hr)) {
                            if (V_VT(&var) == VT_BSTR) {
                                file << _com_util::ConvertBSTRToString(V_BSTR(&var)) << std::endl;
                            }
                            VariantClear(&var);
                        }
                        SysFreeString(bstrName);
                    }
                }
                VariantClear(&var);
            }
        }
        pPropList->Release();
    }
}

//  The EnumeratePropertyValue function
HRESULT EnumeratePropertyValue(IADsPropertyEntry* pEntry, std::ofstream& file)
{
    HRESULT hr = E_FAIL;
    IADsPropertyValue* pValue = NULL;
    IADsLargeInteger* pLargeInt = NULL;
    long lType, lValue;
    BSTR bstr, szString;
    VARIANT var;
    CHAR* pszBOOL = NULL;
    FILETIME filetime;
    SYSTEMTIME systemtime;
    IDispatch* pDisp = NULL;
    DATE date;

    //  For Octet Strings
    void HUGEP* pArray = NULL;
    ULONG dwSLBound;
    ULONG dwSUBound;

    VariantInit(&var);
    hr = pEntry->get_Values(&var);
    if (SUCCEEDED(hr))
    {
        //  Should be a safe array that contains variants
        if (var.vt == (VT_VARIANT | VT_ARRAY))
        {
            VARIANT* pVar;
            long lLBound, lUBound;

            hr = SafeArrayAccessData((SAFEARRAY*)(var.pparray), (void HUGEP * FAR*) & pVar);

            //  One-dimensional array. Get the bounds for the array.
            hr = SafeArrayGetLBound((SAFEARRAY*)(var.pparray), 1, &lLBound);
            hr = SafeArrayGetUBound((SAFEARRAY*)(var.pparray), 1, &lUBound);

            //  Get the count of elements.
            long cElements = lUBound - lLBound + 1;

            //  Get the array elements.
            if (SUCCEEDED(hr))
            {
                for (int i = 0; i < cElements; i++)
                {
                    switch (pVar[i].vt)
                    {
                    case VT_BSTR:
                        wprintf(L"%s ", pVar[i].bstrVal);
                        file << "BSTR: ";
                        file << _com_util::ConvertBSTRToString(V_BSTR(pVar)) << " ";
                        break;

                    case VT_DISPATCH:
                        hr = V_DISPATCH(&pVar[i])->QueryInterface(IID_IADsPropertyValue, (void**)&pValue);
                        if (SUCCEEDED(hr))
                        {
                            hr = pValue->get_ADsType(&lType);
                            switch (lType)
                            {
                            case ADSTYPE_DN_STRING:
                                hr = pValue->get_DNString(&bstr);
                                wprintf(L"%s ", bstr);
                                file << "DN_STRING: ";
                                file << _com_util::ConvertBSTRToString(bstr) << " ";
                                SysFreeString(bstr);
                                break;

                            case ADSTYPE_CASE_EXACT_STRING:
                                hr = pValue->get_CaseExactString(&bstr);
                                wprintf(L"%s ", bstr);
                                file << "CASE_EXACT_STRING: ";
                                file << _com_util::ConvertBSTRToString(bstr) << " ";
                                SysFreeString(bstr);
                                break;

                            case ADSTYPE_CASE_IGNORE_STRING:
                                hr = pValue->get_CaseIgnoreString(&bstr);
                                wprintf(L"%s ", bstr);
                                file << "CASE_IGNORE_STRING: ";
                                file << _com_util::ConvertBSTRToString(bstr) << " ";
                                SysFreeString(bstr);
                                break;

                            case ADSTYPE_NUMERIC_STRING:
                                hr = pValue->get_NumericString(&bstr);
                                wprintf(L"%s ", bstr);
                                file << "NUMERIC_STRING: ";
                                file << _com_util::ConvertBSTRToString(bstr) << " ";
                                SysFreeString(bstr);
                                break;

                            case ADSTYPE_PRINTABLE_STRING:
                                hr = pValue->get_PrintableString(&bstr);
                                wprintf(L"%s ", bstr);
                                file << "PRINTABLE_STRING: ";
                                file << _com_util::ConvertBSTRToString(bstr) << " ";
                                SysFreeString(bstr);
                                break;

                            case ADSTYPE_BOOLEAN:
                                hr = pValue->get_Boolean(&lValue);
                                pszBOOL = (CHAR*)((lValue) ? "TRUE" : "FALSE");
                                printf("%s ", pszBOOL);
                                file << "BOOLEAN: ";
                                file << pszBOOL << " ";
                                break;

                            case ADSTYPE_INTEGER:
                                hr = pValue->get_Integer(&lValue);
                                wprintf(L"%u ", lValue);
                                file << "INTEGER: ";
                                //file << to_string((unsigned long)lValue) << " ";
                                file << to_string(lValue) << " ";
                                break;

                            case ADSTYPE_OCTET_STRING:
                            {
                                VARIANT varOS;

                                VariantInit(&varOS);

                                //  Get the name of the property to handle
                                //  the required properties.
                                pEntry->get_Name(&szString);
                                hr = pValue->get_OctetString(&varOS);

                                //  Get a pointer to the bytes in the octet string.
                                if (SUCCEEDED(hr))
                                {
                                    hr = SafeArrayGetLBound(V_ARRAY(&varOS),
                                        1,
                                        (long FAR*) & dwSLBound);

                                    hr = SafeArrayGetUBound(V_ARRAY(&varOS),
                                        1,
                                        (long FAR*) & dwSUBound);

                                    if (SUCCEEDED(hr))
                                    {
                                        hr = SafeArrayAccessData(V_ARRAY(&varOS), &pArray);
                                    }

                                    file << "OCTET_STRING: ";
                                    if (0 == wcscmp(L"objectGUID", szString))
                                    {
                                        LPOLESTR szDSGUID = new WCHAR[39];

                                        //  Cast to LPGUID
                                        LPGUID pObjectGUID = (LPGUID)pArray;

                                        //  Convert GUID to string.
                                        ::StringFromGUID2(*pObjectGUID, szDSGUID, 39);

                                        //  Print the GUID
                                        wprintf(L"%s ", szDSGUID);
                                        wstring wszGUID(szDSGUID);
                                        string szGUID(wszGUID.begin(), wszGUID.end());
                                        file << szGUID << " ";
                                    }
                                    else if (0 == wcscmp(L"objectSid", szString))
                                    {
                                        PSID pObjectSID = (PSID)pArray;
                                        //  Convert SID to string.
                                        LPOLESTR szSID = NULL;
                                        ConvertSidToStringSid(pObjectSID, &szSID);
                                        wprintf(L"%s ", szSID);
                                        wstring wszSID(szSID);
                                        string szszSID(wszSID.begin(), wszSID.end());
                                        file << szszSID << " ";
                                        LocalFree(szSID);
                                    }
                                    else if (0 == wcscmp(L"cACertificate", szString) ||
                                        0 == wcscmp(L"pKIKeyUsage", szString) ||
                                        0 == wcscmp(L"pKIOverlapPeriod", szString) ||
                                        0 == wcscmp(L"pKIExpirationPeriod", szString) ||
                                        0 == wcscmp(L"authorityRevocationList", szString) ||
                                        0 == wcscmp(L"certificateRevocationList", szString) ||
                                        0 == wcscmp(L"deltaRevocationList", szString) ||
                                        0 == wcscmp(L"userCertificate", szString))
                                    {
                                        for (int i = dwSLBound; i <= dwSUBound; i++)
                                        {
                                            file << std::hex << std::setw(2) << std::setfill('0') << static_cast<int>(((BYTE*)pArray)[i]) << " ";
                                        }
                                    }
                                    else
                                    {
                                        wprintf(L"Value of type Octet String. No Conversion.");
                                        file << "Value of type Octet String. No Conversion.";
                                    }
                                    SafeArrayUnaccessData(V_ARRAY(&varOS));
                                    VariantClear(&varOS);
                                }

                                SysFreeString(szString);
                            }
                            break;

                            case ADSTYPE_UTC_TIME:
                                //  wprintf(L"Value of type UTC_TIME\n");
                                hr = pValue->get_UTCTime(&date);
                                if (SUCCEEDED(hr))
                                {
                                    VARIANT varDate;

                                    //  Pack in variant.vt
                                    varDate.vt = VT_DATE;
                                    varDate.date = date;

                                    VariantChangeType(&varDate, &varDate, VARIANT_NOVALUEPROP, VT_BSTR);
                                    wprintf(L"%s ", varDate.bstrVal);
                                    file << "UTC_TIME: " << std::dec << "(" << (unsigned int)date << ") ";
                                    file << _com_util::ConvertBSTRToString(varDate.bstrVal) << " ";
                                    VariantClear(&varDate);
                                }
                                break;

                            case ADSTYPE_LARGE_INTEGER:
                                //  wprintf(L"Value of type Large Integer\n");
                                //  Get the name of the property to handle
                                //  the required properties.
                                pEntry->get_Name(&szString);
                                hr = pValue->get_LargeInteger(&pDisp);
                                if (SUCCEEDED(hr))
                                {
                                    hr = pDisp->QueryInterface(IID_IADsLargeInteger, (void**)&pLargeInt);
                                    if (SUCCEEDED(hr))
                                    {
                                        hr = pLargeInt->get_HighPart((long*)&filetime.dwHighDateTime);
                                        hr = pLargeInt->get_LowPart((long*)&filetime.dwLowDateTime);
                                        if ((filetime.dwHighDateTime == 0) && (filetime.dwLowDateTime == 0))
                                        {
                                            wprintf(L"No Value ");
                                            file << "No Value ";
                                        }
                                        else
                                        {
                                            //  Verify properties of type LargeInteger that represent time
                                            //  if TRUE, then convert to variant time.
                                            if ((0 == wcscmp(L"accountExpires", szString)) ||
                                                (0 == wcscmp(L"badPasswordTime", szString)) ||
                                                (0 == wcscmp(L"lastLogon", szString)) ||
                                                (0 == wcscmp(L"lastLogoff", szString)) ||
                                                (0 == wcscmp(L"lockoutTime", szString)) ||
                                                (0 == wcscmp(L"pwdLastSet", szString))
                                                )
                                            {
                                                //  Handle special case for Never Expires where low part is -1.
                                                if (filetime.dwLowDateTime == -1)
                                                {
                                                    wprintf(L"Never Expires ");
                                                    file << "Never Expires ";
                                                }
                                                else
                                                {
                                                    if (FileTimeToLocalFileTime(&filetime, &filetime) != 0)
                                                    {
                                                        if (FileTimeToSystemTime(&filetime, &systemtime) != 0)
                                                        {
                                                            if (SystemTimeToVariantTime(&systemtime, &date) != 0)
                                                            {
                                                                VARIANT varDate;

                                                                //  Pack in variant.vt
                                                                varDate.vt = VT_DATE;
                                                                varDate.date = date;

                                                                VariantChangeType(&varDate, &varDate, VARIANT_NOVALUEPROP, VT_BSTR);

                                                                wprintf(L"%s ", varDate.bstrVal);
                                                                file << "LARGE_INTEGER: " << "(" << date << ") ";
                                                                file << _com_util::ConvertBSTRToString(varDate.bstrVal) << " ";

                                                                VariantClear(&varDate);
                                                            }
                                                            else
                                                            {
                                                                wprintf(L"FileTimeToVariantTime failed ");
                                                                file << "FileTimeToVariantTime failed ";
                                                            }
                                                        }
                                                        else
                                                        {
                                                            wprintf(L"FileTimeToSystemTime failed ");
                                                            file << "FileTimeToSystemTime failed ";
                                                        }

                                                    }
                                                    else
                                                    {
                                                        wprintf(L"FileTimeToLocalFileTime failed ");
                                                        file << "FileTimeToLocalFileTime failed ";
                                                    }
                                                }
                                            }
                                            //  Print the LargeInteger.
                                            else
                                            {
                                                wprintf(L"Large Integer: high: %d low: %d ", filetime.dwHighDateTime, filetime.dwLowDateTime);
                                                file << "LARGE_INTEGER: ";
                                                file << "Large Integer: high: " << to_string(filetime.dwHighDateTime) << " low: " << to_string(filetime.dwLowDateTime) << " ";
                                            }
                                        }
                                    }
                                    if (pLargeInt)
                                        pLargeInt->Release();
                                }
                                else
                                {
                                    wprintf(L"Cannot get Large Integer");
                                    file << "Cannot get Large Integer";
                                }

                                if (pDisp)
                                    pDisp->Release();

                                break;

                            case ADSTYPE_NT_SECURITY_DESCRIPTOR:
                                wprintf(L"Value of type NT Security Descriptor ");
                                file << "Value of type NT Security Descriptor ";
                                break;

                            case ADSTYPE_PROV_SPECIFIC:
                                wprintf(L"Value of type Provider Specific ");
                                file << "Value of type Provider Specific ";
                                break;

                            default:
                                wprintf(L"Unhandled ADSTYPE for property value: %d ", lType);
                                file << "Unhandled ADSTYPE for property value: " << to_string(lType) << " ";
                                break;
                            }
                        }
                        else
                        {
                            wprintf(L"QueryInterface failed for IADsPropertyValue. HR: %x\n", hr);
                        }

                        if (pValue)
                        {
                            pValue->Release();
                        }
                        break;

                    default:
                        wprintf(L"Unhandled Variant type for property value array: %d\n", pVar[i].vt);
                        file << "Unhandled Variant type for property value array: " << to_string(pVar[i].vt) << endl;
                        break;
                    }
                }
                wprintf(L"\n");
                file << endl;
            }

            //  Decrement the access count for the array.
            SafeArrayUnaccessData((SAFEARRAY*)(var.pparray));
        }

        VariantClear(&var);
    }

    return hr;
}

//  The GetUserProperties function gets a property list for a user.
HRESULT SaveObjectAttributes(IADs* pObj, std::ofstream& file)
{
    HRESULT hr = E_FAIL;
    LPOLESTR szDSPath = new OLECHAR[MAX_PATH];
    long lCount = 0L;
    long lCountTotal = 0L;
    long lPType = 0L;

    if (!pObj)
    {
        return E_INVALIDARG;
    }

    //  Call GetInfo to load all object properties into the cache
    //  because IADsPropertyList methods read from the cache.
    hr = pObj->GetInfo();
    if (SUCCEEDED(hr))
    {
        IADsPropertyList* pObjProps = NULL;

        //  QueryInterface for an IADsPropertyList pointer.
        hr = pObj->QueryInterface(IID_IADsPropertyList, (void**)&pObjProps);
        if (SUCCEEDED(hr))
        {
            VARIANT var;

            //  Enumerate the properties of the object.
            hr = pObjProps->get_PropertyCount(&lCountTotal);
            wprintf(L"Property Count: %d\n", lCountTotal);

            VariantInit(&var);
            hr = pObjProps->Next(&var);
            if (SUCCEEDED(hr))
            {
                lCount = 1L;
                while (hr == S_OK)
                {
                    if (var.vt == VT_DISPATCH)
                    {
                        IADsPropertyEntry* pEntry = NULL;

                        hr = V_DISPATCH(&var)->QueryInterface(IID_IADsPropertyEntry, (void**)&pEntry);
                        if (SUCCEEDED(hr))
                        {
                            BSTR bstrString;

                            hr = pEntry->get_Name(&bstrString);
                            wprintf(L"%s: ", bstrString);
                            file << _com_util::ConvertBSTRToString(bstrString) << ": ";
                            SysFreeString(bstrString);

                            hr = pEntry->get_ADsType(&lPType);
                            if (lPType != ADSTYPE_INVALID)
                            {
                                hr = EnumeratePropertyValue(pEntry, file);
                                if (FAILED(hr))
                                {
                                    printf("EnumeratePropertyValue failed. hr: %x\n", hr);
                                }
                            }
                            else
                            {
                                wprintf(L"Invalid type\n");
                            }
                        }
                        else
                        {
                            printf("IADsPropertyEntry QueryInterface call failed. hr: %x\n", hr);
                        }

                        // Cleanup.
                        if (pEntry)
                        {
                            pEntry->Release();
                        }
                    }
                    else
                    {
                        printf("Unexpected returned type for VARIANT: %d", var.vt);
                    }
                    VariantClear(&var);
                    hr = pObjProps->Next(&var);
                    if (SUCCEEDED(hr))
                    {
                        lCount++;
                    }
                }
            }
            //  Cleanup.
            pObjProps->Release();
        }

        wprintf(L"Total properties retrieved: %d\n", lCount);
    }

    //  Return S_OK if all properties were retrieved.
    if (lCountTotal == lCount)
    {
        hr = S_OK;
    }

    return hr;
}

void EnumerateADsObjects(IADsContainer* pADsContainer, std::ofstream& file) {
    IEnumVARIANT* pEnum = NULL;
    HRESULT hr;
    VARIANT var;

    hr = ADsBuildEnumerator(pADsContainer, &pEnum);
    if (SUCCEEDED(hr)) {
        ULONG lFetch = 0;
        while (SUCCEEDED(ADsEnumerateNext(pEnum, 1, &var, &lFetch)) && (lFetch > 0)) {
            if (V_VT(&var) == VT_DISPATCH) {
                IADs* pADs = (IADs*)V_DISPATCH(&var);
                //IADs* pADs = NULL;
                //hr = V_DISPATCH(&var)->QueryInterface(IID_IADs, (void**)&pADs);
                BSTR bstrName;
                pADs->get_Name(&bstrName);
                file << "Object: " << _com_util::ConvertBSTRToString(bstrName) << std::endl;
                BSTR bstrADsPath;
                pADs->get_ADsPath(&bstrADsPath);
                file << "ADsPath: " << _com_util::ConvertBSTRToString(bstrADsPath) << std::endl;
                SaveObjectAttributes(pADs, file);

                wprintf(L"\n");
                file << endl;
                // Check if object is indeed a container
                IADsContainer* pChildContainer = nullptr;
                hr = pADs->QueryInterface(IID_IADsContainer, (void**)&pChildContainer);
                if (SUCCEEDED(hr) && pChildContainer) {
                    EnumerateADsObjects(pChildContainer, file);
                }
                pChildContainer->Release();

                pADs->Release();
                SysFreeString(bstrName);
            }
            VariantClear(&var);
        }
        if (pEnum)
            ADsFreeEnumerator(pEnum);
    }
}

// https://learn.microsoft.com/en-us/windows/win32/adsi/enumeration-helper-functions
#define MAX_ENUM      100
HRESULT TestEnumObject(IADsContainer* pADsContainer, std::ofstream& file)
{
    ULONG cElementFetched = 0L;
    IEnumVARIANT* pEnumVariant = NULL;
    VARIANT VariantArray[MAX_ENUM];
    HRESULT hr = S_OK;
    DWORD dwObjects = 0, dwEnumCount = 0, i = 0;
    BOOL  fContinue = TRUE;

    hr = ADsBuildEnumerator(pADsContainer, &pEnumVariant);
    if (FAILED(hr))
    {
        printf("ADsBuildEnumerator failed with %lx\n", hr);
        goto exitpoint;
    }

    fContinue = TRUE;
    while (fContinue) {

        IADs* pObject;
        hr = ADsEnumerateNext(pEnumVariant, MAX_ENUM, VariantArray, &cElementFetched);
        if (FAILED(hr))
        {
            printf("ADsEnumerateNext failed with %lx\n", hr);
            goto exitpoint;
        }

        if (hr == S_FALSE) {
            fContinue = FALSE;
        }

        dwEnumCount++;

        for (i = 0; i < cElementFetched; i++) {
            //IDispatch* pDispatch = NULL;
            BSTR        bstrADsPath = NULL;

            //pDispatch = VariantArray[i].pdispVal;
            hr = V_DISPATCH(VariantArray + i)->QueryInterface(IID_IADs, (void**)&pObject);
            if (SUCCEEDED(hr))
            {
                pObject->get_ADsPath(&bstrADsPath);
                printf("%S\n", bstrADsPath);
                BSTR bstrName;
                pObject->get_Name(&bstrName);
                file << "Object: " << _com_util::ConvertBSTRToString(bstrName) << std::endl;
                SaveObjectAttributes(pObject, file);
                // Check if object is indeed a container
                IADsContainer* pChildContainer = nullptr;
                hr = pObject->QueryInterface(IID_IADsContainer, (void**)&pChildContainer);
                if (SUCCEEDED(hr) && pChildContainer) {
                    TestEnumObject(pChildContainer, file);
                }
            }

            pObject->Release();
            VariantClear(VariantArray + i);
            SysFreeString(bstrADsPath);
        }

        dwObjects += cElementFetched;
    }
    if (dwObjects)
        printf("Total Number of Objects enumerated is %d\n", dwObjects);

exitpoint:
    if (pEnumVariant) {
        ADsFreeEnumerator(pEnumVariant);
    }

    return(hr);
}

void tokenize(const std::wstring& s, const wchar_t delim, std::vector<std::wstring>& out)
{
    std::wstring::size_type beg = 0;
    for (auto end = 0; (end = s.find(delim, end)) != std::wstring::npos; ++end)
    {
        out.push_back(s.substr(beg, end - beg));
        beg = end + 1;
    }

    out.push_back(s.substr(beg));
}

int wmain(int argc, wchar_t* argv[]) {

    if (argc != 3)
    {
        wcout << argv[0] << " -domain " << "<domain name eg. Win2019Dom.com> -> Enumerate Public Key Service" << endl;
        wcout << argv[0] << " -ldap " << "<LDAP name eg. LDAP://CN=Public Key Services,CN=Services,CN=Configuration,DC=Win209Dom,DC=com> -> Enumerate any container path" << endl;
        return 0;
    }

    CoInitialize(NULL);

    IADsContainer* pADsContainer;
    //WCHAR szADsPath[] = { L"LDAP://CN=ForestUpdates,CN=Configuration,DC=Win209Dom,DC=com" };
    //WCHAR szADsPath[] = { L"LDAP://CN=Partitions,CN=Configuration,DC=Win209Dom,DC=com" };
    //WCHAR szADsPath[] = { L"LDAP://CN=Public Key Services,CN=Services,CN=Configuration,DC=Win209Dom,DC=com" };
    //WCHAR szADsPath[] = { L"LDAP://CN=Public Key Services,CN=Services,CN=Configuration,DC=secude-swepdm,DC=com" };

    wstring wszLDAPString;
    std::vector<std::wstring> out;
    if (lstrcmpi(argv[1], L"-domain") == 0) {
        wstring s(argv[2]);
        WCHAR delim = L'.';
        tokenize(s, delim, out);
        wszLDAPString = L"LDAP://CN=Public Key Services,CN=Services,CN=Configuration,DC=";
        wszLDAPString += out[0];
        wszLDAPString += L",DC=";
        wszLDAPString += out[1];
    }
    else if (lstrcmpi(argv[1], L"-ldap") == 0) {
        wszLDAPString = argv[2];
    }
    MessageBox(NULL, wszLDAPString.c_str(), L"2", MB_OK);

    HRESULT hr = ADsGetObject(wszLDAPString.c_str(), IID_IADsContainer, (void**)&pADsContainer);
    if (SUCCEEDED(hr)) {
        std::ofstream file("ADObjects.txt");
        if (file.is_open()) {
            EnumerateADsObjects(pADsContainer, file);
            file.close();
        }
        pADsContainer->Release();
    }
    else {
        std::cerr << "Failed to connect to Active Directory" << std::endl;
    }

    //std::ofstream file("ADObjects.txt");
    //if (file.is_open()) {
    //    IADsContainer* pADsContainer = NULL;
    //    WCHAR szADsPath[] = { L"LDAP://CN=ForestUpdates,CN=Configuration,DC=Win209Dom,DC=com" };
    //    HRESULT hr = ADsGetObject((LPWSTR)szADsPath, IID_IADsContainer, (void**)&pADsContainer);
    //    if (FAILED(hr)) {

    //        printf("\"%S\" is not a valid container object.\n", szADsPath);
    //        goto exitpoint0;
    //    }

    //    TestEnumObject(pADsContainer, file);

    //    file.close();

    //    if (pADsContainer) {
    //        pADsContainer->Release();
    //    }
    //}

exitpoint0:
    CoUninitialize();
    return 0;
}

//// https://learn.microsoft.com/en-us/previous-versions/w40768et(v=vs.140)
//#include <windows.h>
//#include <ole2.h>
//#include <math.h>
//#include <wchar.h>
//#include <objbase.h>
//#include <activeds.h>
//#include <atlbase.h>
//
//#include <sddl.h>
//
//#pragma comment(lib, "activeds.lib")
//#pragma comment(lib, "Adsiid.lib")
//
////  The EnumeratePropertyValue function
//HRESULT EnumeratePropertyValue(IADsPropertyEntry* pEntry)
//{
//    HRESULT hr = E_FAIL;
//    IADsPropertyValue* pValue = NULL;
//    IADsLargeInteger* pLargeInt = NULL;
//    long lType, lValue;
//    BSTR bstr, szString;
//    VARIANT var;
//    CHAR* pszBOOL = NULL;
//    FILETIME filetime;
//    SYSTEMTIME systemtime;
//    IDispatch* pDisp = NULL;
//    DATE date;
//
//    //  For Octet Strings
//    void HUGEP* pArray = NULL;
//    ULONG dwSLBound;
//    ULONG dwSUBound;
//
//    VariantInit(&var);
//    hr = pEntry->get_Values(&var);
//    if (SUCCEEDED(hr))
//    {
//        //  Should be a safe array that contains variants
//        if (var.vt == (VT_VARIANT | VT_ARRAY))
//        {
//            VARIANT* pVar;
//            long lLBound, lUBound;
//
//            hr = SafeArrayAccessData((SAFEARRAY*)(var.pparray), (void HUGEP * FAR*) & pVar);
//
//            //  One-dimensional array. Get the bounds for the array.
//            hr = SafeArrayGetLBound((SAFEARRAY*)(var.pparray), 1, &lLBound);
//            hr = SafeArrayGetUBound((SAFEARRAY*)(var.pparray), 1, &lUBound);
//
//            //  Get the count of elements.
//            long cElements = lUBound - lLBound + 1;
//
//            //  Get the array elements.
//            if (SUCCEEDED(hr))
//            {
//                for (int i = 0; i < cElements; i++)
//                {
//                    switch (pVar[i].vt)
//                    {
//                    case VT_BSTR:
//                        wprintf(L"%s ", pVar[i].bstrVal);
//                        break;
//
//                    case VT_DISPATCH:
//                        hr = V_DISPATCH(&pVar[i])->QueryInterface(IID_IADsPropertyValue, (void**)&pValue);
//                        if (SUCCEEDED(hr))
//                        {
//                            hr = pValue->get_ADsType(&lType);
//                            switch (lType)
//                            {
//                            case ADSTYPE_DN_STRING:
//                                hr = pValue->get_DNString(&bstr);
//                                wprintf(L"%s ", bstr);
//                                SysFreeString(bstr);
//                                break;
//
//                            case ADSTYPE_CASE_EXACT_STRING:
//                                hr = pValue->get_CaseExactString(&bstr);
//                                wprintf(L"%s ", bstr);
//                                SysFreeString(bstr);
//                                break;
//
//                            case ADSTYPE_CASE_IGNORE_STRING:
//                                hr = pValue->get_CaseIgnoreString(&bstr);
//                                wprintf(L"%s ", bstr);
//                                SysFreeString(bstr);
//                                break;
//
//                            case ADSTYPE_NUMERIC_STRING:
//                                hr = pValue->get_NumericString(&bstr);
//                                wprintf(L"%s ", bstr);
//                                SysFreeString(bstr);
//                                break;
//
//                            case ADSTYPE_PRINTABLE_STRING:
//                                hr = pValue->get_PrintableString(&bstr);
//                                wprintf(L"%s ", bstr);
//                                SysFreeString(bstr);
//                                break;
//
//                            case ADSTYPE_BOOLEAN:
//                                hr = pValue->get_Boolean(&lValue);
//                                pszBOOL = (CHAR*)((lValue) ? "TRUE" : "FALSE");
//                                wprintf(L"%s ", pszBOOL);
//                                break;
//
//                            case ADSTYPE_INTEGER:
//                                hr = pValue->get_Integer(&lValue);
//                                wprintf(L"%d ", lValue);
//                                break;
//
//                            case ADSTYPE_OCTET_STRING:
//                            {
//                                VARIANT varOS;
//
//                                VariantInit(&varOS);
//
//                                //  Get the name of the property to handle
//                                //  the required properties.
//                                pEntry->get_Name(&szString);
//                                hr = pValue->get_OctetString(&varOS);
//
//                                //  Get a pointer to the bytes in the octet string.
//                                if (SUCCEEDED(hr))
//                                {
//                                    hr = SafeArrayGetLBound(V_ARRAY(&varOS),
//                                        1,
//                                        (long FAR*) & dwSLBound);
//
//                                    hr = SafeArrayGetUBound(V_ARRAY(&varOS),
//                                        1,
//                                        (long FAR*) & dwSUBound);
//
//                                    if (SUCCEEDED(hr))
//                                    {
//                                        hr = SafeArrayAccessData(V_ARRAY(&varOS), &pArray);
//                                    }
//
//                                    if (0 == wcscmp(L"objectGUID", szString))
//                                    {
//                                        LPOLESTR szDSGUID = new WCHAR[39];
//
//                                        //  Cast to LPGUID
//                                        LPGUID pObjectGUID = (LPGUID)pArray;
//
//                                        //  Convert GUID to string.
//                                        ::StringFromGUID2(*pObjectGUID, szDSGUID, 39);
//
//                                        //  Print the GUID
//                                        wprintf(L"%s ", szDSGUID);
//                                    }
//                                    else if (0 == wcscmp(L"objectSid", szString))
//                                    {
//                                        PSID pObjectSID = (PSID)pArray;
//                                        //  Convert SID to string.
//                                        LPOLESTR szSID = NULL;
//                                        ConvertSidToStringSid(pObjectSID, &szSID);
//                                        wprintf(L"%s ", szSID);
//                                        LocalFree(szSID);
//                                    }
//                                    else
//                                    {
//                                        wprintf(L"Value of type Octet String. No Conversion.");
//                                    }
//                                    SafeArrayUnaccessData(V_ARRAY(&varOS));
//                                    VariantClear(&varOS);
//                                }
//
//                                SysFreeString(szString);
//                            }
//                            break;
//
//                            case ADSTYPE_UTC_TIME:
//                                //  wprintf(L"Value of type UTC_TIME\n");
//                                hr = pValue->get_UTCTime(&date);
//                                if (SUCCEEDED(hr))
//                                {
//                                    VARIANT varDate;
//
//                                    //  Pack in variant.vt
//                                    varDate.vt = VT_DATE;
//                                    varDate.date = date;
//
//                                    VariantChangeType(&varDate, &varDate, VARIANT_NOVALUEPROP, VT_BSTR);
//                                    wprintf(L"%s ", varDate.bstrVal);
//                                    VariantClear(&varDate);
//                                }
//                                break;
//
//                            case ADSTYPE_LARGE_INTEGER:
//                                //  wprintf(L"Value of type Large Integer\n");
//                                //  Get the name of the property to handle
//                                //  the required properties.
//                                pEntry->get_Name(&szString);
//                                hr = pValue->get_LargeInteger(&pDisp);
//                                if (SUCCEEDED(hr))
//                                {
//                                    hr = pDisp->QueryInterface(IID_IADsLargeInteger, (void**)&pLargeInt);
//                                    if (SUCCEEDED(hr))
//                                    {
//                                        hr = pLargeInt->get_HighPart((long*)&filetime.dwHighDateTime);
//                                        hr = pLargeInt->get_LowPart((long*)&filetime.dwLowDateTime);
//                                        if ((filetime.dwHighDateTime == 0) && (filetime.dwLowDateTime == 0))
//                                        {
//                                            wprintf(L"No Value ");
//                                        }
//                                        else
//                                        {
//                                            //  Verify properties of type LargeInteger that represent time
//                                            //  if TRUE, then convert to variant time.
//                                            if ((0 == wcscmp(L"accountExpires", szString)) ||
//                                                (0 == wcscmp(L"badPasswordTime", szString)) ||
//                                                (0 == wcscmp(L"lastLogon", szString)) ||
//                                                (0 == wcscmp(L"lastLogoff", szString)) ||
//                                                (0 == wcscmp(L"lockoutTime", szString)) ||
//                                                (0 == wcscmp(L"pwdLastSet", szString))
//                                                )
//                                            {
//                                                //  Handle special case for Never Expires where low part is -1.
//                                                if (filetime.dwLowDateTime == -1)
//                                                {
//                                                    wprintf(L"Never Expires ");
//                                                }
//                                                else
//                                                {
//                                                    if (FileTimeToLocalFileTime(&filetime, &filetime) != 0)
//                                                    {
//                                                        if (FileTimeToSystemTime(&filetime, &systemtime) != 0)
//                                                        {
//                                                            if (SystemTimeToVariantTime(&systemtime, &date) != 0)
//                                                            {
//                                                                VARIANT varDate;
//
//                                                                //  Pack in variant.vt
//                                                                varDate.vt = VT_DATE;
//                                                                varDate.date = date;
//
//                                                                VariantChangeType(&varDate, &varDate, VARIANT_NOVALUEPROP, VT_BSTR);
//
//                                                                wprintf(L"%s ", varDate.bstrVal);
//
//                                                                VariantClear(&varDate);
//                                                            }
//                                                            else
//                                                            {
//                                                                wprintf(L"FileTimeToVariantTime failed ");
//                                                            }
//                                                        }
//                                                        else
//                                                        {
//                                                            wprintf(L"FileTimeToSystemTime failed ");
//                                                        }
//
//                                                    }
//                                                    else
//                                                    {
//                                                        wprintf(L"FileTimeToLocalFileTime failed ");
//                                                    }
//                                                }
//                                            }
//                                            //  Print the LargeInteger.
//                                            else
//                                            {
//                                                wprintf(L"Large Integer: high: %d low: %d ", filetime.dwHighDateTime, filetime.dwLowDateTime);
//                                            }
//                                        }
//                                    }
//                                    if (pLargeInt)
//                                        pLargeInt->Release();
//                                }
//                                else
//                                {
//                                    wprintf(L"Cannot get Large Integer");
//                                }
//
//                                if (pDisp)
//                                    pDisp->Release();
//
//                                break;
//
//                            case ADSTYPE_NT_SECURITY_DESCRIPTOR:
//                                wprintf(L"Value of type NT Security Descriptor ");
//                                break;
//
//                            case ADSTYPE_PROV_SPECIFIC:
//                                wprintf(L"Value of type Provider Specific ");
//                                break;
//
//                            default:
//                                wprintf(L"Unhandled ADSTYPE for property value: %d ", lType);
//                                break;
//                            }
//                        }
//                        else
//                        {
//                            wprintf(L"QueryInterface failed for IADsPropertyValue. HR: %x\n", hr);
//                        }
//
//                        if (pValue)
//                        {
//                            pValue->Release();
//                        }
//                        break;
//
//                    default:
//                        wprintf(L"Unhandled Variant type for property value array: %d\n", pVar[i].vt);
//                        break;
//                    }
//                }
//                wprintf(L"\n");
//            }
//
//            //  Decrement the access count for the array.
//            SafeArrayUnaccessData((SAFEARRAY*)(var.pparray));
//        }
//
//        VariantClear(&var);
//    }
//
//    return hr;
//}
//
////  The GetUserProperties function gets a property list for a user.
//HRESULT GetUserProperties(IADs* pObj)
//{
//    HRESULT hr = E_FAIL;
//    LPOLESTR szDSPath = new OLECHAR[MAX_PATH];
//    long lCount = 0L;
//    long lCountTotal = 0L;
//    long lPType = 0L;
//
//    if (!pObj)
//    {
//        return E_INVALIDARG;
//    }
//
//    //  Call GetInfo to load all object properties into the cache
//    //  because IADsPropertyList methods read from the cache.
//    hr = pObj->GetInfo();
//    if (SUCCEEDED(hr))
//    {
//        IADsPropertyList* pObjProps = NULL;
//
//        //  QueryInterface for an IADsPropertyList pointer.
//        hr = pObj->QueryInterface(IID_IADsPropertyList, (void**)&pObjProps);
//        if (SUCCEEDED(hr))
//        {
//            VARIANT var;
//
//            //  Enumerate the properties of the object.
//            hr = pObjProps->get_PropertyCount(&lCountTotal);
//            wprintf(L"Property Count: %d\n", lCountTotal);
//
//            VariantInit(&var);
//            hr = pObjProps->Next(&var);
//            if (SUCCEEDED(hr))
//            {
//                lCount = 1L;
//                while (hr == S_OK)
//                {
//                    if (var.vt == VT_DISPATCH)
//                    {
//                        IADsPropertyEntry* pEntry = NULL;
//
//                        hr = V_DISPATCH(&var)->QueryInterface(IID_IADsPropertyEntry, (void**)&pEntry);
//                        if (SUCCEEDED(hr))
//                        {
//                            BSTR bstrString;
//
//                            hr = pEntry->get_Name(&bstrString);
//                            wprintf(L"%s: ", bstrString);
//                            SysFreeString(bstrString);
//
//                            hr = pEntry->get_ADsType(&lPType);
//                            if (lPType != ADSTYPE_INVALID)
//                            {
//                                hr = EnumeratePropertyValue(pEntry);
//                                if (FAILED(hr))
//                                {
//                                    printf("EnumeratePropertyValue failed. hr: %x\n", hr);
//                                }
//                            }
//                            else
//                            {
//                                wprintf(L"Invalid type\n");
//                            }
//                        }
//                        else
//                        {
//                            printf("IADsPropertyEntry QueryInterface call failed. hr: %x\n", hr);
//                        }
//
//                        // Cleanup.
//                        if (pEntry)
//                        {
//                            pEntry->Release();
//                        }
//                    }
//                    else
//                    {
//                        printf("Unexpected returned type for VARIANT: %d", var.vt);
//                    }
//                    VariantClear(&var);
//                    hr = pObjProps->Next(&var);
//                    if (SUCCEEDED(hr))
//                    {
//                        lCount++;
//                    }
//                }
//            }
//            //  Cleanup.
//            pObjProps->Release();
//        }
//
//        wprintf(L"Total properties retrieved: %d\n", lCount);
//    }
//
//    //  Return S_OK if all properties were retrieved.
//    if (lCountTotal == lCount)
//    {
//        hr = S_OK;
//    }
//
//    return hr;
//}
//
////  The FindUserByName function searches for users.
//#define NUM_ATTRIBUTES 1
//
//HRESULT FindUserByName(IDirectorySearch* pSearchBase, //  Container to search.
//    LPOLESTR szFindUser, //  Name of user to find.
//    IADs** ppUser) //  Return a pointer to the user.
//{
//    if ((!pSearchBase) || (!szFindUser))
//    {
//        return E_INVALIDARG;
//    }
//
//    HRESULT hrObj = E_FAIL;
//    HRESULT hr = E_FAIL;
//    ADS_SEARCHPREF_INFO SearchPrefs;
//    //  COL for iterations
//    ADS_SEARCH_COLUMN col;
//    //  Handle used for searching
//    ADS_SEARCH_HANDLE hSearch;
//
//    //  Search entire subtree from root.
//    SearchPrefs.dwSearchPref = ADS_SEARCHPREF_SEARCH_SCOPE;
//    SearchPrefs.vValue.dwType = ADSTYPE_INTEGER;
//    SearchPrefs.vValue.Integer = ADS_SCOPE_SUBTREE;
//
//    //  Set the search preference.
//    DWORD dwNumPrefs = 1;
//    hr = pSearchBase->SetSearchPreference(&SearchPrefs, dwNumPrefs);
//    if (FAILED(hr))
//    {
//        return hr;
//    }
//
//    //  Create search filter.
//    LPWSTR pszFormat = (LPWSTR)L"(&(objectCategory=person)(objectClass=user)(cn=%s))";
//    DWORD dwSearchFilterSize = wcslen(pszFormat) + wcslen(szFindUser) + 1;
//    LPWSTR pszSearchFilter = new WCHAR[dwSearchFilterSize];
//    if (NULL == pszSearchFilter)
//    {
//        return E_OUTOFMEMORY;
//    }
//
//    swprintf_s(pszSearchFilter, dwSearchFilterSize, pszFormat, szFindUser);
//
//    //  Set attributes to return.
//    LPWSTR pszAttribute[NUM_ATTRIBUTES] = { (LPWSTR)L"ADsPath" };
//
//    //  Execute the search.
//    hr = pSearchBase->ExecuteSearch(pszSearchFilter,
//        pszAttribute,
//        NUM_ATTRIBUTES,
//        &hSearch);
//    if (SUCCEEDED(hr))
//    {
//        //  Call IDirectorySearch::GetNextRow() to retrieve the next row of data.
//        while (pSearchBase->GetNextRow(hSearch) != S_ADS_NOMORE_ROWS)
//        {
//            //  Loop through the array of passed column names and
//            //  print the data for each column.
//            for (DWORD x = 0; x < NUM_ATTRIBUTES; x++)
//            {
//                //  Get the data for this column.
//                hr = pSearchBase->GetColumn(hSearch, pszAttribute[x], &col);
//                if (SUCCEEDED(hr))
//                {
//                    //  Print the data for the column and free the column.
//                    //  Be aware that the requested attribute is type CaseIgnoreString.
//                    if (ADSTYPE_CASE_IGNORE_STRING == col.dwADsType)
//                    {
//                        hr = ADsOpenObject(col.pADsValues->CaseIgnoreString,
//                            NULL,
//                            NULL,
//                            ADS_SECURE_AUTHENTICATION,
//                            IID_IADs,
//                            (void**)ppUser);
//                        if (SUCCEEDED(hr))
//                        {
//                            wprintf(L"Found User.\n", szFindUser);
//                            wprintf(L"%s: %s\r\n", pszAttribute[x], col.pADsValues->CaseIgnoreString);
//                            hrObj = S_OK;
//                        }
//                    }
//
//                    pSearchBase->FreeColumn(&col);
//                }
//                else
//                {
//                    hr = E_FAIL;
//                }
//            }
//        }
//        //  Close the search handle to cleanup.
//        pSearchBase->CloseSearchHandle(hSearch);
//    }
//
//    delete pszSearchFilter;
//
//    if (FAILED(hrObj))
//    {
//        hr = hrObj;
//    }
//
//    return hr;
//}
//
////  Entry point for the application.
//#define BUFFER_SIZE (MAX_PATH * 2)
//
//void wmain(int argc, wchar_t* argv[])
//{
//    MessageBox(NULL, L"1", L"2", MB_OK);
//
//    //  Handle the command line arguments.
//    WCHAR szBuffer[BUFFER_SIZE];
//    if (argv[1] == NULL)
//    {
//        wprintf(L"This program finds a user in the current Windows 2000 domain\n");
//        wprintf(L"and displays its properties.\n");
//        wprintf(L"Enter Common Name of the user to find:");
//        //fgetws(szBuffer, BUFFER_SIZE, stdin);
//        wscanf_s(L"%s", szBuffer, BUFFER_SIZE);
//    }
//    else
//    {
//        wcsncpy_s(szBuffer, argv[1], BUFFER_SIZE);
//    }
//
//    //  If the string is empty, then exit.
//    if (0 == wcscmp(L"", szBuffer))
//    {
//        return;
//    }
//
//    wprintf(L"\nFinding user: %s...\n", szBuffer);
//
//    //  Initialize COM.
//    CoInitialize(NULL);
//    HRESULT hr = S_OK;
//
//    //  Get rootDSE and the domain container DN.
//    IADs* pObject = NULL;
//    IDirectorySearch* pDS = NULL;
//    LPOLESTR szPath = new OLECHAR[MAX_PATH];
//    hr = ADsOpenObject(L"LDAP://rootDSE",
//        NULL,
//        NULL,
//        ADS_SECURE_AUTHENTICATION, //  Use Secure Authentication.
//        IID_IADs,
//        (void**)&pObject);
//    if (SUCCEEDED(hr))
//    {
//        VARIANT var;
//        VariantInit(&var);
//        hr = pObject->Get(CComBSTR(L"defaultNamingContext"), &var);
//        if (SUCCEEDED(hr))
//        {
//
//            wcscpy_s(szPath, MAX_PATH, L"LDAP://");
//            wcscat_s(szPath, MAX_PATH, var.bstrVal);
//            VariantClear(&var);
//
//            if (pObject)
//            {
//                pObject->Release();
//                pObject = NULL;
//            }
//
//            //  Bind to the root of the current domain.
//            hr = ADsOpenObject(szPath,
//                NULL,
//                NULL,
//                ADS_SECURE_AUTHENTICATION, //  Use Secure Authentication.
//                IID_IDirectorySearch,
//                (void**)&pDS);
//            if (SUCCEEDED(hr))
//            {
//                hr = FindUserByName(pDS, //  Container to search
//                    szBuffer,   //  Name of user to find
//                    &pObject); //  Return a pointer to the user
//                if (SUCCEEDED(hr))
//                {
//                    wprintf(L"----------------------------------------------\n");
//                    wprintf(L"--------Call GetUserProperties-----------\n");
//                    hr = GetUserProperties(pObject);
//                    wprintf(L"GetUserProperties HR: %x\n", hr);
//                }
//                else
//                {
//                    wprintf(L"User \"%s\" not Found.\n", szBuffer);
//                    wprintf(L"FindUserByName failed with the following HR: %x\n", hr);
//                }
//
//                if (pDS)
//                    pDS->Release();
//            }
//
//            if (pObject)
//                pObject->Release();
//        }
//    }
//
//    //  Uninitialize COM.
//    CoUninitialize();
//
//    return;
//}