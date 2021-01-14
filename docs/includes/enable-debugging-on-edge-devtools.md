Lorsque le complément s’exécute dans Microsoft Edge, le code sans interface utilisateur ne peut pas être joint au débogueur par défaut.
Le code sans interface utilisateur est tout code en cours d'exécution lorsque le volet des tâches n'est pas visible, tel que les commandes de complément. Pour activer le débogage, exécutez les commandes [Windows PowerShell](/powershell/scripting/getting-started/getting-started-with-windows-powershell) suivantes :

1. Exécutez la commande suivante pour obtenir des informations sur le package de l’application **Microsoft. Win32WebViewHost**.
    
    ```powershell
    Get-AppxPackage Microsoft.Win32WebViewHost
    ```
    
    La commande répertorie les informations relatives au package de l’application similaires à la sortie suivante.
    
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
    
2. Exécutez la commande suivante pour activer le débogage. Utilisez la valeur de **PackageFullName** répertoriée à partir de la commande précédente.
    
    ```powershell
    setx JS_DEBUG <PackageFullName>
    ```
    
3. Si Office était déjà en cours d’exécution, fermez et redémarrez Office pour qu’il récupère la modification de débogage.