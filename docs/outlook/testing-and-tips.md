---
title: Déployer et installer des compléments Outlook à des fins de test
description: Créez un fichier manifeste, déployez le fichier IU de complément, installez le complément dans votre boîte aux lettres, puis testez-le.
ms.date: 11/06/2019
localization_priority: Priority
ms.openlocfilehash: 521199a87282b58c3bf10553886174e8be26cacf
ms.sourcegitcommit: a3ddfdb8a95477850148c4177e20e56a8673517c
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 02/20/2020
ms.locfileid: "42166074"
---
# <a name="deploy-and-install-outlook-add-ins-for-testing"></a><span data-ttu-id="700ef-103">Déployer et installer des compléments Outlook à des fins de test</span><span class="sxs-lookup"><span data-stu-id="700ef-103">Deploy and install Outlook add-ins for testing</span></span>

<span data-ttu-id="700ef-104">Dans le cadre du processus de développement d’un complément Outlook, vous devrez déployer et installer de façon itérative le complément à des fins de test, ce qui implique les étapes suivantes :</span><span class="sxs-lookup"><span data-stu-id="700ef-104">As part of the process of developing an Outlook add-in, you will probably find yourself iteratively deploying and installing the add-in for testing, which involves the following steps:</span></span>

1. <span data-ttu-id="700ef-105">Création d’un fichier manifeste qui décrit le complément.</span><span class="sxs-lookup"><span data-stu-id="700ef-105">Creating a manifest file that describes the add-in.</span></span>
1. <span data-ttu-id="700ef-106">Déploiement du ou des fichiers de l’interface utilisateur du complément sur un serveur web.</span><span class="sxs-lookup"><span data-stu-id="700ef-106">Deploying the add-in UI file(s) to a web server.</span></span>
1. <span data-ttu-id="700ef-107">Installation du complément dans votre boîte aux lettres.</span><span class="sxs-lookup"><span data-stu-id="700ef-107">Installing the add-in in your mailbox.</span></span>
1. <span data-ttu-id="700ef-108">Test du complément, mise en œuvre des modifications appropriées dans l’interface utilisateur ou dans les fichiers manifeste, et répétition des étapes 2 et 3 pour tester les modifications.</span><span class="sxs-lookup"><span data-stu-id="700ef-108">Testing the add-in, making appropriate changes to the UI or manifest files, and repeating steps 2 and 3 to test the changes.</span></span>

> [!NOTE]
> <span data-ttu-id="700ef-109">[Les volets personnalisés sont déconseillés](https://developer.microsoft.com/outlook/blogs/make-your-add-ins-available-in-the-office-ribbon/) afin de vous assurer que vous utilisez [un point d’extension de complément pris en charge](outlook-add-ins-overview.md#extension-points).</span><span class="sxs-lookup"><span data-stu-id="700ef-109">[Custom panes have been deprecated](https://developer.microsoft.com/outlook/blogs/make-your-add-ins-available-in-the-office-ribbon/) so please ensure that you're using [a supported add-in extension point](outlook-add-ins-overview.md#extension-points).</span></span>

## <a name="create-a-manifest-file-for-the-add-in"></a><span data-ttu-id="700ef-110">Création d’un fichier manifeste pour le complément</span><span class="sxs-lookup"><span data-stu-id="700ef-110">Create a manifest file for the add-in</span></span>

<span data-ttu-id="700ef-p101">Chaque complément est décrit par un manifeste XML, un document qui fournit au serveur des informations sur le complément, décrit le complément pour l’utilisateur et identifie l’emplacement du fichier HTML de l’interface utilisateur du complément. Vous pouvez stocker le manifeste dans un dossier local ou sur un serveur, à condition que le complément soit accessible par le serveur Exchange de la boîte aux lettres avec laquelle vous procédez aux tests. Nous partons du principe que vous stockez votre manifeste dans un dossier local. Pour plus d’informations sur la création d’un fichier manifeste, voir [Manifestes des compléments Outlook](manifests.md).</span><span class="sxs-lookup"><span data-stu-id="700ef-p101">Each add-in is described by an XML manifest, a document that gives the server information about the add-in, provides descriptive information about the add-in for the user, and identifies the location of the add-in UI HTML file. You can store the manifest in a local folder or server, as long as the location is accessible by the Exchange server of the mailbox that you are testing with. We'll assume that you store your manifest in a local folder. For information about how to create a manifest file, see [Outlook add-in manifests](manifests.md).</span></span>

## <a name="deploy-an-add-in-to-a-web-server"></a><span data-ttu-id="700ef-115">Déploiement d’un complément sur un serveur web</span><span class="sxs-lookup"><span data-stu-id="700ef-115">Deploy an add-in to a web server</span></span>

<span data-ttu-id="700ef-p102">Vous pouvez utiliser du code HTML et JavaScript pour créer le complément. Les fichiers source obtenus sont stockés sur un serveur web accessible par le biais du serveur Exchange qui héberge le complément. Après le déploiement initial des fichiers source pour le complément, vous pouvez mettre à jour l’interface utilisateur et le comportement du complément en remplaçant les fichiers HTML ou JavaScript stocké sur le serveur web par une nouvelle version du fichier HTML.</span><span class="sxs-lookup"><span data-stu-id="700ef-p102">You can use HTML and JavaScript to create the add-in. The resulting source files are stored on a web server that can be accessed by the Exchange server that hosts the add-in. After initially deploying the source files for the add-in, you can update the add-in UI and behavior by replacing the HTML files or JavaScript files stored on the web server with a new version of the HTML file.</span></span>

## <a name="install-the-add-in"></a><span data-ttu-id="700ef-119">Installer le complément</span><span class="sxs-lookup"><span data-stu-id="700ef-119">Install the add-in</span></span>

<span data-ttu-id="700ef-120">Après la préparation du fichier manifeste du complément et le déploiement de son interface utilisateur sur un serveur web accessible, vous pouvez charger une version test du complément pour une boîte aux lettres sur un serveur Exchange à l’aide d’un client Outlook ou installer le complément en exécutant des cmdlets Windows PowerShell à distance.</span><span class="sxs-lookup"><span data-stu-id="700ef-120">After preparing the add-in manifest file and deploying the add-in UI to a web server that can be accessed, you can sideload the add-in for a mailbox on an Exchange server by using an Outlook client, or install the add-in by running remote Windows PowerShell cmdlets.</span></span>

### <a name="sideload-the-add-in"></a><span data-ttu-id="700ef-121">Charger une version test du complément</span><span class="sxs-lookup"><span data-stu-id="700ef-121">Sideload the add-in</span></span>

<span data-ttu-id="700ef-p103">Vous pouvez installer un complément si votre boîte aux lettres est sur Exchange Online, Exchange 2013 ou une version ultérieure. Les compléments de chargement de version test nécessitent au minimum le rôle **Mes compléments personnalisés** pour votre serveur Exchange. Pour tester votre complément ou installer des compléments en général en spécifiant une URL ou un nom de fichier pour le manifeste de complément, vous devez demander à votre administrateur Exchange de vous octroyer les autorisations nécessaires.</span><span class="sxs-lookup"><span data-stu-id="700ef-p103">You can install an add-in if your mailbox is on Exchange Online, Exchange 2013 or a later release. Sideloading add-ins requires at minimum the **My Custom Apps** role for your Exchange Server. In order to test your add-in, or install add-ins in general by specifying a URL or file name for the add-in manifest, you should request your Exchange administrator to provide the necessary permissions.</span></span>

<span data-ttu-id="700ef-p104">L’administrateur Exchange peut exécuter la cmdlet PowerShell suivante pour affecter les autorisations nécessaires à un seul utilisateur. Dans cet exemple, `wendyri` est l’alias de messagerie de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="700ef-p104">The Exchange administrator can run the following PowerShell cmdlet to assign a single user the necessary permissions. In this example, `wendyri` is the user's email alias.</span></span>

```powershell
New-ManagementRoleAssignment -Role "My Custom Apps" -User "wendyri"
```

<span data-ttu-id="700ef-127">Selon les besoins, l’administrateur peut exécuter la cmdlet suivante pour affecter des autorisations nécessaires similaires à plusieurs utilisateurs :</span><span class="sxs-lookup"><span data-stu-id="700ef-127">If necessary, the administrator can run the following cmdlet to assign multiple users the similar necessary permissions:</span></span>

```powershell
$users = Get-Mailbox *$users | ForEach-Object { New-ManagementRoleAssignment -Role "My Custom Apps" -User $_.Alias}
```

<span data-ttu-id="700ef-128">Pour plus d’informations sur le rôle « Mes compléments personnalisés », consultez la rubrique relative au [rôle « Mes compléments personnalisés »](/exchange/my-custom-apps-role-exchange-2013-help).</span><span class="sxs-lookup"><span data-stu-id="700ef-128">For more information about the My Custom Apps role, see [My Custom Apps role](/exchange/my-custom-apps-role-exchange-2013-help).</span></span>

<span data-ttu-id="700ef-129">L’utilisation d’Office 365 ou de Visual Studio pour développer des compléments vous amène à endosser le rôle d’administrateur d’organisation, ce qui vous permet d’installer des compléments par fichier ou par URL dans le Centre d’administration Exchange ou via des cmdlets PowerShell.</span><span class="sxs-lookup"><span data-stu-id="700ef-129">Using Office 365 or Visual Studio to develop add-ins assigns you the organization administrator role which allows you to install add-ins by file or URL in the EAC, or by Powershell cmdlets.</span></span>

### <a name="install-an-add-in-by-using-remote-powershell"></a><span data-ttu-id="700ef-130">Installation d’un complément à l’aide de PowerShell à distance</span><span class="sxs-lookup"><span data-stu-id="700ef-130">Install an add-in by using remote PowerShell</span></span>

<span data-ttu-id="700ef-131">Après avoir créé une session Windows PowerShell à distance sur votre serveur Exchange, vous pouvez installer un complément Outlook en utilisant la cmdlet `New-App` avec la commande PowerShell suivante.</span><span class="sxs-lookup"><span data-stu-id="700ef-131">After you create a remote Windows PowerShell session on your Exchange server, you can install an Outlook add-in by using the `New-App` cmdlet with the following PowerShell command.</span></span>

```powershell
New-App -URL:"http://<fully-qualified URL">
```

<span data-ttu-id="700ef-132">L’URL complète est l’emplacement du fichier de manifeste de complément que vous avez préparé pour votre complément.</span><span class="sxs-lookup"><span data-stu-id="700ef-132">The fully qualified URL is the location of the add-in manifest file that you prepared for your add-in.</span></span>

<span data-ttu-id="700ef-133">Vous pouvez utiliser les cmdlets supplémentaires suivantes pour gérer les compléments pour une boîte aux lettres :</span><span class="sxs-lookup"><span data-stu-id="700ef-133">You can use the following additional PowerShell cmdlets to manage the add-ins for a mailbox:</span></span>

-  <span data-ttu-id="700ef-134">`Get-App` : répertorie les compléments activés pour une boîte aux lettres.</span><span class="sxs-lookup"><span data-stu-id="700ef-134">`Get-App` - Lists the add-ins that are enabled for a mailbox.</span></span>
-  <span data-ttu-id="700ef-135">`Set-App` : active ou désactive un complément sur une boîte aux lettres.</span><span class="sxs-lookup"><span data-stu-id="700ef-135">`Set-App` - Enables or disables a add-in on a mailbox.</span></span>
-  <span data-ttu-id="700ef-136">`Remove-App` : supprime un complément précédemment installé à partir d’un serveur Exchange.</span><span class="sxs-lookup"><span data-stu-id="700ef-136">`Remove-App` - Removes a previously installed add-in from an Exchange server.</span></span>

## <a name="client-versions"></a><span data-ttu-id="700ef-137">Versions client</span><span class="sxs-lookup"><span data-stu-id="700ef-137">Client versions</span></span>

<span data-ttu-id="700ef-138">Le choix des versions du client Outlook à tester dépend de vos besoins en matière de développement.</span><span class="sxs-lookup"><span data-stu-id="700ef-138">Deciding what versions of the Outlook client to test depends on your development requirements.</span></span>

- <span data-ttu-id="700ef-p105">Si vous développez un complément pour une utilisation privée ou uniquement pour les membres de votre organisation, il est important de tester les versions d’Outlook que votre entreprise utilise. Gardez à l’esprit que certains utilisateurs peuvent utiliser Outlook sur le web. Par conséquent, vous devez également tester les versions des navigateurs standard utilisés au sein de votre entreprise.</span><span class="sxs-lookup"><span data-stu-id="700ef-p105">If you are developing an add-in for private use, or only for members of your organization, then it is important to test the versions of Outlook that your company uses. Keep in mind that some users may use Outlook on the web, so testing your company's standard browser versions is also important.</span></span>

- <span data-ttu-id="700ef-p106">Si vous développez un complément pour [AppSource](https://appsource.microsoft.com), vous devez tester les versions requises tel que spécifié dans les [stratégies de validation d’AppSource 4.12.1](/office/dev/store/validation-policies#4-apps-and-add-ins-behave-predictably). Cela inclut notamment :</span><span class="sxs-lookup"><span data-stu-id="700ef-p106">If you are developing an add-in to list in [AppSource](https://appsource.microsoft.com), you must test the required versions as specified in the [AppSource validation policies 4.12.1](/office/dev/store/validation-policies#4-apps-and-add-ins-behave-predictably). This includes:</span></span>
    - <span data-ttu-id="700ef-143">la dernière et avant-dernière version d’Outlook sur Windows ;</span><span class="sxs-lookup"><span data-stu-id="700ef-143">The latest version of Outlook on Windows and the version prior to the latest.</span></span>
    - <span data-ttu-id="700ef-144">la dernière version d’Outlook sur Mac ;</span><span class="sxs-lookup"><span data-stu-id="700ef-144">The latest version of Outlook on Mac.</span></span>
    - <span data-ttu-id="700ef-145">la dernière version d’Outlook sur iOS et Android (si votre complément [prend en charge le facteur de forme pour mobile](add-mobile-support.md)) ;</span><span class="sxs-lookup"><span data-stu-id="700ef-145">The latest version of Outlook on iOS and Android (if your add-in [supports mobile form factor](add-mobile-support.md)).</span></span>
    - <span data-ttu-id="700ef-146">les versions de navigateur spécifiées dans la stratégie de validation d’AppSource 4.12.1.</span><span class="sxs-lookup"><span data-stu-id="700ef-146">The browser versions specified in AppSource validation policy 4.12.1.</span></span>

> [!NOTE]
> <span data-ttu-id="700ef-147">Si votre complément ne prend pas en charge l’un des clients ci-dessus car il demande [un ensemble de conditions requises d’API](apis.md) que le client ne prend pas en charge, ce client est supprimé de la liste des clients requis.</span><span class="sxs-lookup"><span data-stu-id="700ef-147">If your add-in does not support one of the above clients due to [requesting an API requirement set](apis.md) that the client does not support, that client would be removed from the list of required clients.</span></span>

## <a name="see-also"></a><span data-ttu-id="700ef-148">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="700ef-148">See also</span></span>

- [<span data-ttu-id="700ef-149">Résolution des erreurs rencontrées par l’utilisateur avec des compléments Office</span><span class="sxs-lookup"><span data-stu-id="700ef-149">Troubleshoot user errors with Office Add-ins</span></span>](../testing/testing-and-troubleshooting.md)
