<span data-ttu-id="b455d-101">Lorsque le complément s’exécute dans Microsoft Edge, le code sans interface utilisateur ne peut pas être joint au débogueur par défaut.</span><span class="sxs-lookup"><span data-stu-id="b455d-101">When the add-in is running in Microsoft Edge, UI-less code will not be able to attach to a debugger by default.</span></span>
<span data-ttu-id="b455d-102">Le code sans interface utilisateur est tout code en cours d'exécution lorsque le volet des tâches n'est pas visible, tel que les commandes de complément.</span><span class="sxs-lookup"><span data-stu-id="b455d-102">UI-less code is any code running while the task pane is not visible, such as add-in commands.</span></span> <span data-ttu-id="b455d-103">Pour activer le débogage, exécutez les commandes [Windows PowerShell](/powershell/scripting/getting-started/getting-started-with-windows-powershell) suivantes :</span><span class="sxs-lookup"><span data-stu-id="b455d-103">To enable debugging, you need to run the following [Windows PowerShell](/powershell/scripting/getting-started/getting-started-with-windows-powershell) commands.</span></span>

1. <span data-ttu-id="b455d-104">Exécutez la commande suivante pour obtenir des informations sur le package de l’application **Microsoft. Win32WebViewHost**.</span><span class="sxs-lookup"><span data-stu-id="b455d-104">Run the following command to get information for the **Microsoft.Win32WebViewHost** app package.</span></span>
    
    ```powershell
    Get-AppxPackage Microsoft.Win32WebViewHost
    ```
    
    <span data-ttu-id="b455d-105">La commande répertorie les informations relatives au package de l’application similaires à la sortie suivante.</span><span class="sxs-lookup"><span data-stu-id="b455d-105">The command lists app package information similar to the following output.</span></span>
    
    ```powershell
    Name              : Microsoft.Win32WebViewHost
    Publisher         : CN=Microsoft Windows, O=Microsoft Corporation, L=Redmond, S=Washington, C=US
    Architecture      : Neutral
    ResourceId        : neutral
    Version           : 10.0.18362.449
    PackageFullName   : Microsoft.Win32WebViewHost_10.0.18362.449_neutral_neutral_cw5n1h2txyewy
    InstallLocation   : C:\Windows\SystemApps\Microsoft.Win32WebViewHost_cw5n1h2txyewy
    IsFramework       : False
    PackageFamilyName : Microsoft.Win32WebViewHost_cw5n1h2txyewy
    PublisherId       : cw5n1h2txyewy
    IsResourcePackage : False
    IsBundle          : False
    IsDevelopmentMode : False
    NonRemovable      : True
    IsPartiallyStaged : False
    SignatureKind     : System
    Status            : Ok
    ```
    
2. <span data-ttu-id="b455d-106">Exécutez la commande suivante pour activer le débogage.</span><span class="sxs-lookup"><span data-stu-id="b455d-106">Run the following command to enable debugging.</span></span> <span data-ttu-id="b455d-107">Utilisez la valeur de **PackageFullName** répertoriée à partir de la commande précédente.</span><span class="sxs-lookup"><span data-stu-id="b455d-107">Use the value for the **PackageFullName** listed from the previous command.</span></span>
    
    ```powershell
    setx JS_DEBUG <PackageFullName>
    ```
    
3. <span data-ttu-id="b455d-108">Si Office était déjà en cours d’exécution, fermez et redémarrez Office pour qu’il récupère la modification de débogage.</span><span class="sxs-lookup"><span data-stu-id="b455d-108">If Office was already running, close and restart Office so that it picks up the debugging change.</span></span>