---
title: Valider et résoudre des problèmes avec votre manifeste
description: Utiliser ces méthodes pour valider le manifeste des compléments Office.
ms.date: 08/15/2019
localization_priority: Priority
ms.openlocfilehash: bf70aca68135073ed92d2e4d2c176b944836c7ad
ms.sourcegitcommit: da8e6148f4bd9884ab9702db3033273a383d15f0
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/20/2019
ms.locfileid: "36477921"
---
# <a name="validate-and-troubleshoot-issues-with-your-manifest"></a><span data-ttu-id="26297-103">Valider et résoudre des problèmes avec votre manifeste</span><span class="sxs-lookup"><span data-stu-id="26297-103">Validate and troubleshoot issues with your manifest</span></span>

<span data-ttu-id="26297-104">Vous souhaitez peut-être valider le fichier manifeste de votre complément pour vous assurer que celui-ci est correct et complet.</span><span class="sxs-lookup"><span data-stu-id="26297-104">You may want to validate your add-in's manifest file to ensure that it's correct and complete.</span></span> <span data-ttu-id="26297-105">La validation peut également identifier les problèmes à l’origine de l’erreur « Votre manifeste de complément n’est pas valide » lorsque vous essayez de charger une version test de votre complément.</span><span class="sxs-lookup"><span data-stu-id="26297-105">Validation can also identify issues that are causing the error "Your add-in manifest is not valid" when you attempt to sideload your add-in.</span></span> <span data-ttu-id="26297-106">Cet article décrit plusieurs méthodes de validation du fichier manifeste et de résolution des problèmes liés à votre complément.</span><span class="sxs-lookup"><span data-stu-id="26297-106">This article describes multiple ways to validate the manifest file and troubleshoot problems with your add-in.</span></span>

## <a name="validate-your-manifest-with-the-yeoman-generator-for-office-add-ins"></a><span data-ttu-id="26297-107">Valider votre manifeste avec le générateur Yeoman pour les compléments Office</span><span class="sxs-lookup"><span data-stu-id="26297-107">Validate your manifest with the Yeoman generator for Office Add-ins</span></span>

<span data-ttu-id="26297-108">Si vous avez utilisé [le générateur Yeoman pour les compléments Office](https://www.npmjs.com/package/generator-office) pour créer votre complément, vous pouvez également l’utiliser pour valider le fichier manifeste de votre projet.</span><span class="sxs-lookup"><span data-stu-id="26297-108">If you used the [Yeoman generator for Office Add-ins](https://www.npmjs.com/package/generator-office) to create your add-in, you can also use it to validate your project's manifest file.</span></span> <span data-ttu-id="26297-109">Exécutez la commande suivante dans le répertoire racine de votre projet :</span><span class="sxs-lookup"><span data-stu-id="26297-109">Run the following command in the root directory of your project:</span></span>

```command&nbsp;line
npm run validate
```

![Gif animé qui montre le validateur Yo Office exécuté sur la ligne de commande et les résultats générés indiquant « Validation Passed » (validation réussie)](../images/yo-office-validator.gif)

> [!NOTE]
> <span data-ttu-id="26297-111">Pour accéder à cette fonctionnalité, votre projet de complément doit être créé à l’aide du [générateur Yeoman pour les compléments Office](https://www.npmjs.com/package/generator-office) (version 1.1.17 ou ultérieure).</span><span class="sxs-lookup"><span data-stu-id="26297-111">To have access to this functionality, your add-in project must have been created by using [Yeoman generator for Office Add-ins](https://www.npmjs.com/package/generator-office) version 1.1.17 or later.</span></span>

## <a name="validate-your-manifest-with-office-addin-manifest"></a><span data-ttu-id="26297-112">Valider votre manifeste avec office-addin-manifest</span><span class="sxs-lookup"><span data-stu-id="26297-112">Validate your manifest with office-addin-manifest</span></span>

<span data-ttu-id="26297-113">Si vous n’avez pas utilisé [le générateur Yeoman pour les compléments Office](https://www.npmjs.com/package/generator-office) pour créer votre complément, vous pouvez valider le fichier manifeste à l’aide de [office-addin-manifest](https://www.npmjs.com/package/office-addin-manifest).</span><span class="sxs-lookup"><span data-stu-id="26297-113">If you didn't use the [Yeoman generator for Office Add-ins](https://www.npmjs.com/package/generator-office) to create your add-in, you can validate the manifest by using [office-addin-manifest](https://www.npmjs.com/package/office-addin-manifest).</span></span>

1. <span data-ttu-id="26297-114">Installez [Node.js](https://nodejs.org/download/).</span><span class="sxs-lookup"><span data-stu-id="26297-114">Install [Node.js](https://nodejs.org/download/).</span></span>

2. <span data-ttu-id="26297-115">Exécutez la commande suivante dans le répertoire racine de votre projet.</span><span class="sxs-lookup"><span data-stu-id="26297-115">Run the following command in the root directory of your project.</span></span> <span data-ttu-id="26297-116">Remplacez `MANIFEST_FILE` par le nom du fichier manifeste.</span><span class="sxs-lookup"><span data-stu-id="26297-116">Replace `MANIFEST_FILE` with the name of the manifest file.</span></span>

    ```command&nbsp;line
    npx office-addin-manifest validate MANIFEST_FILE
    ```

    > [!NOTE]
    > <span data-ttu-id="26297-117">Si elle s’exécute, la commande renvoie le message d’erreur « La syntaxe de la commande n’est pas valide »</span><span class="sxs-lookup"><span data-stu-id="26297-117">If running this command results in the error message "The command syntax is not valid."</span></span> <span data-ttu-id="26297-118">(étant donné que la commande `validate` n’est pas reconnue), exécutez la commande suivante pour valider le manifeste (en remplaçant `MANIFEST_FILE` par le nom du fichier manifeste) :</span><span class="sxs-lookup"><span data-stu-id="26297-118">(because the `validate` command is not recognized), run the following command to validate the manifest (replacing `MANIFEST_FILE` with the name of the manifest file):</span></span> 
    > 
    > `npx --ignore-existing office-addin-manifest validate MANIFEST_FILE`

## <a name="validate-your-manifest-against-the-xml-schema"></a><span data-ttu-id="26297-119">Validez votre manifeste par rapport au schéma XML</span><span class="sxs-lookup"><span data-stu-id="26297-119">Validate your manifest against the XML schema</span></span>

<span data-ttu-id="26297-120">Vous pouvez valider le fichier manifeste par rapport aux fichiers de [définition de schéma XML (XSD)](https://github.com/OfficeDev/office-js-docs-pr/tree/master/docs/overview/schemas).</span><span class="sxs-lookup"><span data-stu-id="26297-120">You can validate a manifest against the [XML Schema Definition (XSD)](https://github.com/OfficeDev/office-js-docs-pr/tree/master/docs/overview/schemas) files.</span></span> <span data-ttu-id="26297-121">Cela permet de s’assurer que le fichier manifeste suit le schéma approprié, y compris les espaces de noms pour les éléments que vous utilisez.</span><span class="sxs-lookup"><span data-stu-id="26297-121">To help ensure that the manifest file follows the correct schema, including any namespaces for the elements you are using.</span></span> <span data-ttu-id="26297-122">Si vous avez copié des éléments à partir d’autres exemples de manifestes, vérifiez par deux fois que vous avez également **inclus les espaces de noms appropriés**.</span><span class="sxs-lookup"><span data-stu-id="26297-122">If you copied elements from other sample manifests double check you also **include the appropriate namespaces**.</span></span> <span data-ttu-id="26297-123">Pour ce faire, vous pouvez utiliser un outil de validation de schéma XML.</span><span class="sxs-lookup"><span data-stu-id="26297-123">You can use an XML schema validation tool to perform this validation.</span></span>

### <a name="to-use-a-command-line-xml-schema-validation-tool-to-validate-your-manifest"></a><span data-ttu-id="26297-124">Pour utiliser un outil de validation de schéma XML à ligne de commande pour valider votre manifeste</span><span class="sxs-lookup"><span data-stu-id="26297-124">To use a command-line XML schema validation tool to validate your manifest</span></span>

1. <span data-ttu-id="26297-125">Installez [tar](https://www.gnu.org/software/tar/) et [libxml](http://xmlsoft.org/FAQ.html), si vous ne l’avez pas déjà fait.</span><span class="sxs-lookup"><span data-stu-id="26297-125">Install [tar](https://www.gnu.org/software/tar/) and [libxml](http://xmlsoft.org/FAQ.html), if you haven't already.</span></span>

2. <span data-ttu-id="26297-p106">Exécutez la commande suivante. Remplacez `XSD_FILE` par le chemin d’accès au fichier XSD manifeste et `XML_FILE` par le chemin d’accès au fichier XML manifeste.</span><span class="sxs-lookup"><span data-stu-id="26297-p106">Run the following command. Replace `XSD_FILE` with the path to the manifest XSD file, and replace `XML_FILE` with the path to the manifest XML file.</span></span>
    
    ```command&nbsp;line
    xmllint --noout --schema XSD_FILE XML_FILE
    ```

## <a name="use-runtime-logging-to-debug-your-add-in"></a><span data-ttu-id="26297-128">Utilisation de la journalisation runtime pour déboguer votre complément</span><span class="sxs-lookup"><span data-stu-id="26297-128">Use runtime logging to debug your add-in</span></span>

<span data-ttu-id="26297-129">Vous pouvez utiliser la journalisation runtime pour déboguer le manifeste de votre complément ainsi que plusieurs erreurs d’installation.</span><span class="sxs-lookup"><span data-stu-id="26297-129">You can use runtime logging to debug your add-in's manifest as well as several installation errors.</span></span> <span data-ttu-id="26297-130">Cette fonctionnalité peut vous aider à identifier et à résoudre les problèmes avec votre manifeste qui ne sont pas détectés par la validation de schéma XSD, comme une incompatibilité entre les ID de ressources.</span><span class="sxs-lookup"><span data-stu-id="26297-130">This feature can help you identify and fix issues with your manifest that are not detected by XSD schema validation, such as a mismatch between resource IDs.</span></span> <span data-ttu-id="26297-131">La journalisation runtime est particulièrement utile pour déboguer des compléments qui implémentent des commandes de complément et des fonctions personnalisées Excel.</span><span class="sxs-lookup"><span data-stu-id="26297-131">Runtime logging is particularly  useful for debugging add-ins that implement add-in commands and Excel custom functions.</span></span>   

> [!NOTE]
> <span data-ttu-id="26297-132">La fonctionnalité de journalisation runtime est actuellement disponible pour Office 2016 pour ordinateur de bureau.</span><span class="sxs-lookup"><span data-stu-id="26297-132">The runtime logging feature is currently available for Office 2016 desktop.</span></span>

### <a name="to-turn-on-runtime-logging"></a><span data-ttu-id="26297-133">Pour activer la journalisation runtime</span><span class="sxs-lookup"><span data-stu-id="26297-133">To turn on runtime logging</span></span>

> [!IMPORTANT]
> <span data-ttu-id="26297-p108">La journalisation runtime réduit les performances. Activez-la uniquement lorsque vous avez besoin de déboguer des problèmes avec votre manifeste de complément.</span><span class="sxs-lookup"><span data-stu-id="26297-p108">Runtime Logging affects performance. Turn it on only when you need to debug issues with your add-in manifest.</span></span>

<span data-ttu-id="26297-136">Pour activer la journalisation runtime, procédez comme suit :</span><span class="sxs-lookup"><span data-stu-id="26297-136">To turn on runtime logging:</span></span>

1. <span data-ttu-id="26297-137">Vérifiez que vous exécutez la version Bureau d’Office 2016 **16.0.7019** ou une version ultérieure.</span><span class="sxs-lookup"><span data-stu-id="26297-137">Make sure that you are running Office 2016 desktop build **16.0.7019** or later.</span></span> 

2. <span data-ttu-id="26297-138">Ajoutez la clé de registre `RuntimeLogging` sous `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\WEF\Developer\`.</span><span class="sxs-lookup"><span data-stu-id="26297-138">Add the `RuntimeLogging` registry key under `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\WEF\Developer\`.</span></span> 

    > [!NOTE]
    > <span data-ttu-id="26297-139">Si la clé (dossier) `Developer` n’existe pas sous `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\WEF\`, procédez comme suit pour la créer :</span><span class="sxs-lookup"><span data-stu-id="26297-139">If the `Developer` key (folder) does not already exist under `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\WEF\`, complete the following steps to create it:</span></span> 
    > 1. <span data-ttu-id="26297-140">Cliquez avec le bouton droit de votre souris sur la clé (dossier) **WEF**, puis sélectionnez **Nouveau** > **Clé**.</span><span class="sxs-lookup"><span data-stu-id="26297-140">Right-click the **WEF** key (folder) and select **New** > **Key**.</span></span>
    > 2. <span data-ttu-id="26297-141">Nommez la nouvelle clé **Développeur**.</span><span class="sxs-lookup"><span data-stu-id="26297-141">Name the new key **Developer**.</span></span>

3. <span data-ttu-id="26297-p109">Définissez la valeur par défaut de la clé pour le chemin d’accès complet du fichier dans lequel écrire le journal. Pour obtenir un exemple, reportez-vous à [EnableRuntimeLogging.zip](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/raw/master/Tools/RuntimeLogging/EnableRuntimeLogging.zip).</span><span class="sxs-lookup"><span data-stu-id="26297-p109">Set the default value of the key to the full path of the file where you want the log to be written. For an example, see [EnableRuntimeLogging.zip](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/raw/master/Tools/RuntimeLogging/EnableRuntimeLogging.zip).</span></span> 

    > [!NOTE]
    > <span data-ttu-id="26297-144">Le répertoire dans lequel le fichier journal sera écrit doit déjà exister et vous devez disposer des autorisations d’écriture correspondantes.</span><span class="sxs-lookup"><span data-stu-id="26297-144">The directory in which the log file will be written must already exist, and you must have write permissions to it.</span></span> 
 
<span data-ttu-id="26297-p110">L’image suivante indique à quoi doit ressembler le registre. Pour désactiver la fonctionnalité, supprimez la clé de registre `RuntimeLogging`.</span><span class="sxs-lookup"><span data-stu-id="26297-p110">The following image shows what the registry should look like. To turn the feature off, remove the `RuntimeLogging` key from the registry.</span></span> 

![Capture d’écran de l’Éditeur du Registre avec la clé de registre RuntimeLogging](http://i.imgur.com/Sa9TyI6.png)

### <a name="to-troubleshoot-issues-with-your-manifest"></a><span data-ttu-id="26297-148">Résolution des problèmes avec votre manifeste</span><span class="sxs-lookup"><span data-stu-id="26297-148">To troubleshoot issues with your manifest</span></span>

<span data-ttu-id="26297-149">Pour utiliser la journalisation runtime pour résoudre les problèmes de chargement d’un complément, procédez comme suit :</span><span class="sxs-lookup"><span data-stu-id="26297-149">To use runtime logging to troubleshoot issues loading an add-in:</span></span>
 
1. <span data-ttu-id="26297-150">[Chargez une version test de votre complément](sideload-office-add-ins-for-testing.md).</span><span class="sxs-lookup"><span data-stu-id="26297-150">[Sideload your add-in](sideload-office-add-ins-for-testing.md) for testing.</span></span> 

    > [!NOTE]
    > <span data-ttu-id="26297-151">Nous vous recommandons de charger uniquement une version test du complément que vous testez pour réduire le nombre de messages dans le fichier journal.</span><span class="sxs-lookup"><span data-stu-id="26297-151">We recommend that you sideload only the add-in that you are testing to minimize the number of messages in the log file.</span></span>

2. <span data-ttu-id="26297-152">Si rien ne se produit et que votre complément n’apparaît pas (et ne s’affiche pas dans la boîte de dialogue des compléments), ouvrez le fichier journal.</span><span class="sxs-lookup"><span data-stu-id="26297-152">If nothing happens and you don't see your add-in (and it's not appearing in the add-ins dialog box), open the log file.</span></span>

3. <span data-ttu-id="26297-p111">Recherchez le fichier journal pour l’ID de votre complément, que vous définissez dans votre manifeste. Dans le fichier journal, cet ID est intitulé `SolutionId`.</span><span class="sxs-lookup"><span data-stu-id="26297-p111">Search the log file for your add-in ID, which you define in your manifest. In the log file, this ID is labeled `SolutionId`.</span></span> 

<span data-ttu-id="26297-p112">Dans l’exemple suivant, le fichier journal identifie un contrôle qui pointe vers un fichier de ressources qui n’existe pas. Pour cet exemple, la correction consistera à corriger la faute de frappe dans le manifeste ou à ajouter la ressource manquante.</span><span class="sxs-lookup"><span data-stu-id="26297-p112">In the following example, the log file identifies a control that points to a resource file that doesn't exist. For this example, the fix would be to correct the typo in the manifest or to add the missing resource.</span></span>

![Capture d’écran d’un fichier journal avec une entrée qui spécifie un ID de ressource qui est introuvable](http://i.imgur.com/f8bouLA.png) 

### <a name="known-issues-with-runtime-logging"></a><span data-ttu-id="26297-158">Problèmes connus avec la journalisation runtime</span><span class="sxs-lookup"><span data-stu-id="26297-158">Known issues with runtime logging</span></span>

<span data-ttu-id="26297-p113">Vous pouvez afficher des messages dans le fichier journal qui sont source de confusion ou classés de façon incorrecte. Par exemple :</span><span class="sxs-lookup"><span data-stu-id="26297-p113">You might see messages in the log file that are confusing or that are classified incorrectly. For example:</span></span>

- <span data-ttu-id="26297-161">Le message `Medium Current host not in add-in's host list` suivi de `Unexpected Parsed manifest targeting different host` est classé incorrectement en tant qu’erreur.</span><span class="sxs-lookup"><span data-stu-id="26297-161">The message `Medium Current host not in add-in's host list` followed by `Unexpected Parsed manifest targeting different host` is incorrectly classified as an error.</span></span>

- <span data-ttu-id="26297-162">Si vous voyez le message `Unexpected Add-in is missing required manifest fields DisplayName` et qu’il ne contient pas de SolutionId, l’erreur n’est probablement pas liée au complément que vous déboguez.</span><span class="sxs-lookup"><span data-stu-id="26297-162">If you see the message `Unexpected Add-in is missing required manifest fields DisplayName` and it doesn't contain a SolutionId, the error is most likely not related to the add-in you are debugging.</span></span> 

- <span data-ttu-id="26297-p114">Tous les messages `Monitorable` sont des erreurs attendues du point de vue du système. Parfois, ils indiquent un problème avec votre manifeste, comme un élément mal orthographié qui a été ignoré, mais n’a pas provoqué l’échec du manifeste.</span><span class="sxs-lookup"><span data-stu-id="26297-p114">Any `Monitorable` messages are expected errors from a system point of view. Sometimes they indicate an issue with your manifest, such as a misspelled element that was skipped but didn't cause the manifest to fail.</span></span> 

## <a name="clear-the-office-cache"></a><span data-ttu-id="26297-165">Vider le cache Office</span><span class="sxs-lookup"><span data-stu-id="26297-165">Clear the Office cache</span></span>

<span data-ttu-id="26297-166">Si les modifications apportées au manifeste, par exemple aux noms de fichier des icônes de bouton dans le ruban ou au texte des commandes de complément, ne semblent pas être appliquées, essayez de vider le cache Office de votre ordinateur.</span><span class="sxs-lookup"><span data-stu-id="26297-166">If changes you've made in the manifest, such as file names of ribbon button icons, or text of add-in commands, do not seem to take effect, try clearing the Office cache on your computer.</span></span> 

#### <a name="for-windows"></a><span data-ttu-id="26297-167">Pour Windows :</span><span class="sxs-lookup"><span data-stu-id="26297-167">For Windows:</span></span>
<span data-ttu-id="26297-168">Supprimer le contenu du dossier `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\`.</span><span class="sxs-lookup"><span data-stu-id="26297-168">Delete the content of the folder `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\`.</span></span>

#### <a name="for-mac"></a><span data-ttu-id="26297-169">Pour Mac :</span><span class="sxs-lookup"><span data-stu-id="26297-169">For Mac:</span></span>

[!include[additional cache folders on Mac](../includes/mac-cache-folders.md)]

#### <a name="for-ios"></a><span data-ttu-id="26297-170">Pour iOS :</span><span class="sxs-lookup"><span data-stu-id="26297-170">For iOS:</span></span>
<span data-ttu-id="26297-p115">Appelez `window.location.reload(true)` à partir de JavaScript dans le complément pour forcer le rechargement. Vous pouvez également choisir de réinstaller Office.</span><span class="sxs-lookup"><span data-stu-id="26297-p115">Call `window.location.reload(true)` from JavaScript in the add-in to force a reload. Alternatively, you can reinstall Office.</span></span>

## <a name="see-also"></a><span data-ttu-id="26297-173">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="26297-173">See also</span></span>

- [<span data-ttu-id="26297-174">Manifeste XML des compléments Office</span><span class="sxs-lookup"><span data-stu-id="26297-174">Office Add-ins XML manifest</span></span>](../develop/add-in-manifests.md)
- [<span data-ttu-id="26297-175">Chargement de la version test des compléments Office</span><span class="sxs-lookup"><span data-stu-id="26297-175">Sideload Office Add-ins for testing</span></span>](sideload-office-add-ins-for-testing.md)
- [<span data-ttu-id="26297-176">Débogage des compléments Office</span><span class="sxs-lookup"><span data-stu-id="26297-176">Debug Office Add-ins</span></span>](debug-add-ins-using-f12-developer-tools-on-windows-10.md)
