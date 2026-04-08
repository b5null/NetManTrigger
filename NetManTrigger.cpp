// https://stackoverflow.com/questions/5917304/how-do-i-detect-a-disabled-network-interface-connection-from-a-windows-applicati/5942359#5942359

#include <iostream>
#include <comdef.h>
#include <netcon.h>
#include <stdio.h>

const CLSID CLSID_ConnectionManager = {0xBA126AD1, 0x2166, 0x11D1, {0xB1, 0xD0, 0x00, 0x80, 0x5F, 0xC1, 0x27, 0x0E}};
const IID IID_INetConnectionManager = {0xC08956A2, 0x1CD3, 0x11D1, {0xB1, 0xC5, 0x00, 0x80, 0x5F, 0xC1, 0x27, 0x0E}};

int main()
{
    FILE* log = fopen("C:\\Users\\tier1-mgmt\\Documents\\out.txt", "w");
    if (!log) return 1;

    HRESULT hResult;

    typedef void(__stdcall* LPNcFreeNetconProperties)(NETCON_PROPERTIES* pProps);
    HMODULE hModule = LoadLibraryA("netshell.dll");
    if (hModule == NULL) { fprintf(log, "[-] LoadLibrary failed\n"); fclose(log); return 1; }
    fprintf(log, "[+] netshell.dll loaded\n");

    LPNcFreeNetconProperties NcFreeNetconProperties = (LPNcFreeNetconProperties)GetProcAddress(hModule, "NcFreeNetconProperties");
    if (!NcFreeNetconProperties) { fprintf(log, "[-] GetProcAddress failed\n"); fclose(log); return 1; }
    fprintf(log, "[+] NcFreeNetconProperties resolved\n");

    hResult = CoInitializeEx(0, COINIT_MULTITHREADED);
    if (SUCCEEDED(hResult))
    {
        fprintf(log, "[+] CoInitializeEx OK\n");
        INetConnectionManager* pConnectionManager = 0;
        hResult = CoCreateInstance(CLSID_ConnectionManager, 0, CLSCTX_ALL, IID_INetConnectionManager, (void**)&pConnectionManager);
        if (SUCCEEDED(hResult))
        {
            fprintf(log, "[+] CoCreateInstance OK\n");
            IEnumNetConnection* pEnumConnection = 0;
            hResult = pConnectionManager->EnumConnections(NCME_DEFAULT, &pEnumConnection);
            if (SUCCEEDED(hResult))
            {
                fprintf(log, "[+] EnumConnections OK\n");
                INetConnection* pConnection = 0;
                ULONG count;
                while (pEnumConnection->Next(1, &pConnection, &count) == S_OK)
                {
                    NETCON_PROPERTIES* pConnectionProperties = 0;
                    hResult = pConnection->GetProperties(&pConnectionProperties);
                    if (SUCCEEDED(hResult))
                    {
                        fprintf(log, "[+] Interface: %ls\n", pConnectionProperties->pszwName);
                        NcFreeNetconProperties(pConnectionProperties);
                    }
                    else
                        fprintf(log, "[-] GetProperties failed: 0x%08X\n", hResult);
                    pConnection->Release();
                }
                pEnumConnection->Release();
            }
            else
                fprintf(log, "[-] EnumConnections failed: 0x%08X\n", hResult);
            pConnectionManager->Release();
        }
        else
            fprintf(log, "[-] CoCreateInstance failed: 0x%08X\n", hResult);
        CoUninitialize();
    }
    else
        fprintf(log, "[-] CoInitializeEx failed: 0x%08X\n", hResult);

    FreeLibrary(hModule);
    fprintf(log, "[+] Done\n");
    fclose(log);
    return 0;
}
