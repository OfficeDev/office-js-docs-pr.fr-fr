---
title: Valider et résoudre des problèmes avec votre manifeste
description: Utiliser ces méthodes pour valider le manifeste des compléments Office.
ms.date: 09/18/2019
localization_priority: Priority
ms.openlocfilehash: c320c05b944bba9e24a4d3c0e5ef514ac13cc3c6
ms.sourcegitcommit: a0257feabcfe665061c14b8bdb70cf82f7aca414
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/18/2019
ms.locfileid: "37035335"
---
# <a name="validate-and-troubleshoot-issues-with-your-manifest"></a><span data-ttu-id="65a33-103">Valider et résoudre des problèmes avec votre manifeste</span><span class="sxs-lookup"><span data-stu-id="65a33-103">Validate and troubleshoot issues with your manifest</span></span>

<span data-ttu-id="65a33-104">Vous souhaitez peut-être valider le fichier manifeste de votre complément pour vous assurer que celui-ci est correct et complet.</span><span class="sxs-lookup"><span data-stu-id="65a33-104">You may want to validate your add-in's manifest file to ensure that it's correct and complete.</span></span> <span data-ttu-id="65a33-105">La validation peut également identifier les problèmes à l’origine de l’erreur « Votre manifeste de complément n’est pas valide » lorsque vous essayez de charger une version test de votre complément.</span><span class="sxs-lookup"><span data-stu-id="65a33-105">Validation can also identify issues that are causing the error "Your add-in manifest is not valid" when you attempt to sideload your add-in.</span></span> <span data-ttu-id="65a33-106">Cet article décrit plusieurs méthodes de validation du fichier manifeste et de résolution des problèmes liés à votre complément.</span><span class="sxs-lookup"><span data-stu-id="65a33-106">This article describes multiple ways to validate the manifest file and troubleshoot problems with your add-in.</span></span>

## <a name="validate-your-manifest-with-the-yeoman-generator-for-office-add-ins"></a><span data-ttu-id="65a33-107">Valider votre manifeste avec le générateur Yeoman pour les compléments Office</span><span class="sxs-lookup"><span data-stu-id="65a33-107">Validate your manifest with the Yeoman generator for Office Add-ins</span></span>

<span data-ttu-id="65a33-108">Si vous avez utilisé [le générateur Yeoman pour les compléments Office](https://www.npmjs.com/package/generator-office) pour créer votre complément, vous pouvez également l’utiliser pour valider le fichier manifeste de votre projet.</span><span class="sxs-lookup"><span data-stu-id="65a33-108">If you used the [Yeoman generator for Office Add-ins](https://www.npmjs.com/package/generator-office) to create your add-in, you can also use it to validate your project's manifest file.</span></span> <span data-ttu-id="65a33-109">Exécutez la commande suivante dans le répertoire racine de votre projet :</span><span class="sxs-lookup"><span data-stu-id="65a33-109">Run the following command in the root directory of your project:</span></span>

```command&nbsp;line
npm run validate
```

![Gif animé qui montre le validateur Yo Office exécuté sur la ligne de commande et les résultats générés indiquant « Validation Passed » (validation réussie)](../images/yo-office-validator.gif)

> [!NOTE]
> <span data-ttu-id="65a33-111">Pour accéder à cette fonctionnalité, votre projet de complément doit être créé à l’aide du [générateur Yeoman pour les compléments Office](https://www.npmjs.com/package/generator-office) (version 1.1.17 ou ultérieure).</span><span class="sxs-lookup"><span data-stu-id="65a33-111">To have access to this functionality, your add-in project must have been created by using [Yeoman generator for Office Add-ins](https://www.npmjs.com/package/generator-office) version 1.1.17 or later.</span></span>

## <a name="validate-your-manifest-with-office-addin-manifest"></a><span data-ttu-id="65a33-112">Valider votre manifeste avec office-addin-manifest</span><span class="sxs-lookup"><span data-stu-id="65a33-112">Validate your manifest with office-addin-manifest</span></span>

<span data-ttu-id="65a33-113">Si vous n’avez pas utilisé [le générateur Yeoman pour les compléments Office](https://www.npmjs.com/package/generator-office) pour créer votre complément, vous pouvez valider le fichier manifeste à l’aide de [office-addin-manifest](https://www.npmjs.com/package/office-addin-manifest).</span><span class="sxs-lookup"><span data-stu-id="65a33-113">If you didn't use the [Yeoman generator for Office Add-ins](https://www.npmjs.com/package/generator-office) to create your add-in, you can validate the manifest by using [office-addin-manifest](https://www.npmjs.com/package/office-addin-manifest).</span></span>

1. <span data-ttu-id="65a33-114">Installez [Node.js](https://nodejs.org/download/).</span><span class="sxs-lookup"><span data-stu-id="65a33-114">Install [Node.js](https://nodejs.org/download/).</span></span>

2. <span data-ttu-id="65a33-115">Exécutez la commande suivante dans le répertoire racine de votre projet.</span><span class="sxs-lookup"><span data-stu-id="65a33-115">Run the following command in the root directory of your project.</span></span> <span data-ttu-id="65a33-116">Remplacez `MANIFEST_FILE` par le nom du fichier manifeste.</span><span class="sxs-lookup"><span data-stu-id="65a33-116">Replace `MANIFEST_FILE` with the name of the manifest file.</span></span>

    ```command&nbsp;line
    npx office-addin-manifest validate MANIFEST_FILE
    ```

    > [!NOTE]
    > <span data-ttu-id="65a33-117">Si elle s’exécute, la commande renvoie le message d’erreur « La syntaxe de la commande n’est pas valide »</span><span class="sxs-lookup"><span data-stu-id="65a33-117">If running this command results in the error message "The command syntax is not valid."</span></span> <span data-ttu-id="65a33-118">(étant donné que la commande `validate` n’est pas reconnue), exécutez la commande suivante pour valider le manifeste (en remplaçant `MANIFEST_FILE` par le nom du fichier manifeste) :</span><span class="sxs-lookup"><span data-stu-id="65a33-118">(because the `validate` command is not recognized), run the following command to validate the manifest (replacing `MANIFEST_FILE` with the name of the manifest file):</span></span> 
    > 
    > `npx --ignore-existing office-addin-manifest validate MANIFEST_FILE`

## <a name="validate-your-manifest-against-the-xml-schema"></a><span data-ttu-id="65a33-119">Validez votre manifeste par rapport au schéma XML</span><span class="sxs-lookup"><span data-stu-id="65a33-119">Validate your manifest against the XML schema</span></span>

<span data-ttu-id="65a33-120">Vous pouvez valider le fichier manifeste par rapport aux fichiers de [définition de schéma XML (XSD)](https://github.com/OfficeDev/office-js-docs-pr/tree/master/docs/overview/schemas).</span><span class="sxs-lookup"><span data-stu-id="65a33-120">You can validate a manifest against the [XML Schema Definition (XSD)](https://github.com/OfficeDev/office-js-docs-pr/tree/master/docs/overview/schemas) files.</span></span> <span data-ttu-id="65a33-121">Cela permet de s’assurer que le fichier manifeste suit le schéma approprié, y compris les espaces de noms pour les éléments que vous utilisez.</span><span class="sxs-lookup"><span data-stu-id="65a33-121">To help ensure that the manifest file follows the correct schema, including any namespaces for the elements you are using.</span></span> <span data-ttu-id="65a33-122">Si vous avez copié des éléments à partir d’autres exemples de manifestes, vérifiez par deux fois que vous avez également **inclus les espaces de noms appropriés**.</span><span class="sxs-lookup"><span data-stu-id="65a33-122">If you copied elements from other sample manifests double check you also **include the appropriate namespaces**.</span></span> <span data-ttu-id="65a33-123">Pour ce faire, vous pouvez utiliser un outil de validation de schéma XML.</span><span class="sxs-lookup"><span data-stu-id="65a33-123">You can use an XML schema validation tool to perform this validation.</span></span>

### <a name="to-use-a-command-line-xml-schema-validation-tool-to-validate-your-manifest"></a><span data-ttu-id="65a33-124">Pour utiliser un outil de validation de schéma XML à ligne de commande pour valider votre manifeste</span><span class="sxs-lookup"><span data-stu-id="65a33-124">To use a command-line XML schema validation tool to validate your manifest</span></span>

1. <span data-ttu-id="65a33-125">Installez [tar](https://www.gnu.org/software/tar/) et [libxml](http://xmlsoft.org/FAQ.html), si vous ne l’avez pas déjà fait.</span><span class="sxs-lookup"><span data-stu-id="65a33-125">Install [tar](https://www.gnu.org/software/tar/) and [libxml](http://xmlsoft.org/FAQ.html), if you haven't already.</span></span>

2. <span data-ttu-id="65a33-p106">Exécutez la commande suivante. Remplacez `XSD_FILE` par le chemin d’accès au fichier XSD manifeste et `XML_FILE` par le chemin d’accès au fichier XML manifeste.</span><span class="sxs-lookup"><span data-stu-id="65a33-p106">Run the following command. Replace `XSD_FILE` with the path to the manifest XSD file, and replace `XML_FILE` with the path to the manifest XML file.</span></span>
    
    ```command&nbsp;line
    xmllint --noout --schema XSD_FILE XML_FILE
    ```

## <a name="use-runtime-logging-to-debug-your-add-in"></a><span data-ttu-id="65a33-128">Utilisation de la journalisation runtime pour déboguer votre complément</span><span class="sxs-lookup"><span data-stu-id="65a33-128">Use runtime logging to debug your add-in</span></span>

<span data-ttu-id="65a33-129">Vous pouvez utiliser la journalisation runtime pour déboguer le manifeste de votre complément ainsi que plusieurs erreurs d’installation.</span><span class="sxs-lookup"><span data-stu-id="65a33-129">You can use runtime logging to debug your add-in's manifest as well as several installation errors.</span></span> <span data-ttu-id="65a33-130">Cette fonctionnalité peut vous aider à identifier et à résoudre les problèmes avec votre manifeste qui ne sont pas détectés par la validation de schéma XSD, comme une incompatibilité entre les ID de ressources.</span><span class="sxs-lookup"><span data-stu-id="65a33-130">This feature can help you identify and fix issues with your manifest that are not detected by XSD schema validation, such as a mismatch between resource IDs.</span></span> <span data-ttu-id="65a33-131">La journalisation runtime est particulièrement utile pour déboguer des compléments qui implémentent des commandes de complément et des fonctions personnalisées Excel.</span><span class="sxs-lookup"><span data-stu-id="65a33-131">Runtime logging is particularly  useful for debugging add-ins that implement add-in commands and Excel custom functions.</span></span>   

> [!NOTE]
> <span data-ttu-id="65a33-132">La fonctionnalité de journalisation runtime est actuellement disponible pour Office 2016 pour ordinateur de bureau.</span><span class="sxs-lookup"><span data-stu-id="65a33-132">The runtime logging feature is currently available for Office 2016 desktop.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="65a33-133">La journalisation runtime affecte les performances.</span><span class="sxs-lookup"><span data-stu-id="65a33-133">Runtime Logging affects performance.</span></span> <span data-ttu-id="65a33-134">Activez-la uniquement lorsque vous avez besoin de déboguer des problèmes avec votre manifeste de complément.</span><span class="sxs-lookup"><span data-stu-id="65a33-134">Turn it on only when you need to debug issues with your add-in manifest.</span></span>

### <a name="runtime-logging-on-windows"></a><span data-ttu-id="65a33-135">Journalisation de l’exécution sur Windows</span><span class="sxs-lookup"><span data-stu-id="65a33-135">Runtime logging on Windows</span></span>

1. <span data-ttu-id="65a33-136">Vérifiez que vous exécutez la version Bureau d’Office 2016 **16.0.7019** ou une version ultérieure.</span><span class="sxs-lookup"><span data-stu-id="65a33-136">Make sure that you are running Office 2016 desktop build **16.0.7019** or later.</span></span> 

2. <span data-ttu-id="65a33-137">Ajoutez la clé de registre `RuntimeLogging` sous `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\WEF\Developer\`.</span><span class="sxs-lookup"><span data-stu-id="65a33-137">Add the `RuntimeLogging` registry key under `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\WEF\Developer\`.</span></span> 

    > [!NOTE]
    > <span data-ttu-id="65a33-138">Si la clé (dossier) `Developer` n’existe pas sous `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\WEF\`, procédez comme suit pour la créer :</span><span class="sxs-lookup"><span data-stu-id="65a33-138">If the `Developer` key (folder) does not already exist under `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\WEF\`, complete the following steps to create it:</span></span> 
    > 1. <span data-ttu-id="65a33-139">Cliquez avec le bouton droit de votre souris sur la clé (dossier) **WEF**, puis sélectionnez **Nouveau** > **Clé**.</span><span class="sxs-lookup"><span data-stu-id="65a33-139">Right-click the **WEF** key (folder) and select **New** > **Key**.</span></span>
    > 2. <span data-ttu-id="65a33-140">Nommez la nouvelle clé **Développeur**.</span><span class="sxs-lookup"><span data-stu-id="65a33-140">Name the new key **Developer**.</span></span>

3. <span data-ttu-id="65a33-p109">Définissez la valeur par défaut de la clé pour le chemin d’accès complet du fichier dans lequel écrire le journal. Pour obtenir un exemple, reportez-vous à [EnableRuntimeLogging.zip](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/raw/master/Tools/RuntimeLogging/EnableRuntimeLogging.zip).</span><span class="sxs-lookup"><span data-stu-id="65a33-p109">Set the default value of the key to the full path of the file where you want the log to be written. For an example, see [EnableRuntimeLogging.zip](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/raw/master/Tools/RuntimeLogging/EnableRuntimeLogging.zip).</span></span> 

    > [!NOTE]
    > <span data-ttu-id="65a33-143">Le répertoire dans lequel le fichier journal sera écrit doit déjà exister et vous devez disposer des autorisations d’écriture correspondantes.</span><span class="sxs-lookup"><span data-stu-id="65a33-143">The directory in which the log file will be written must already exist, and you must have write permissions to it.</span></span> 
 
<span data-ttu-id="65a33-p110">L’image suivante indique à quoi doit ressembler le registre. Pour désactiver la fonctionnalité, supprimez la clé de registre `RuntimeLogging`.</span><span class="sxs-lookup"><span data-stu-id="65a33-p110">The following image shows what the registry should look like. To turn the feature off, remove the `RuntimeLogging` key from the registry.</span></span> 

![Capture d’écran de l’Éditeur du Registre avec la clé de registre RuntimeLogging](http://i.imgur.com/Sa9TyI6.png)

### <a name="runtime-logging-on-mac"></a><span data-ttu-id="65a33-147">Journalisation de l’exécution sur Mac</span><span class="sxs-lookup"><span data-stu-id="65a33-147">Runtime logging on Mac</span></span>

1. <span data-ttu-id="65a33-148">Vérifiez que vous exécutez la version Bureau d’Office 2016 **16.27** (19071500) ou une version ultérieure.</span><span class="sxs-lookup"><span data-stu-id="65a33-148">Make sure that you are running Office 2016 desktop build **16.0.7019** or later.</span></span>

2. <span data-ttu-id="65a33-149">Ouvrez **Terminal** et configurez une préférence de journalisation de l’exécution à l’aide de la commande `defaults` :</span><span class="sxs-lookup"><span data-stu-id="65a33-149">Open **Terminal** and set a runtime logging preference by using the `defaults` command:</span></span>
    
    ```command&nbsp;line
    defaults write <bundle id> CEFRuntimeLoggingFile -string <file_name>
    ```

    <span data-ttu-id="65a33-150">`<bundle id>` identifie l’hôte pour lequel activer la journalisation de l’exécution.</span><span class="sxs-lookup"><span data-stu-id="65a33-150">`<bundle id>` identifies which the host for which to enable runtime logging.</span></span> <span data-ttu-id="65a33-151">`<file_name>` est le nom du fichier texte dans lequel le journal sera écrit.</span><span class="sxs-lookup"><span data-stu-id="65a33-151">`<file_name>` is the name of the text file to which the log will be written.</span></span>

    <span data-ttu-id="65a33-152">Configurez `<bundle id>` à l’une des valeurs suivantes pour activer la journalisation de l’exécution pour l’hôte correspondant :</span><span class="sxs-lookup"><span data-stu-id="65a33-152">Set `<bundle id>` to one of the following values to enable runtime logging for the corresponding host:</span></span>

    - `com.microsoft.Word`
    - `com.microsoft.Excel`
    - `com.microsoft.Powerpoint`
    - `com.microsoft.Outlook`

<span data-ttu-id="65a33-153">L’exemple suivant montre comment activer la journalisation de l’exécution pour Word, puis ouvrir le fichier journal :</span><span class="sxs-lookup"><span data-stu-id="65a33-153">The following example enables runtime logging for Word and then opens the log file:</span></span>

```command&nbsp;line
defaults write com.microsoft.Word CEFRuntimeLoggingFile -string "runtime_logs.txt"
open ~/library/Containers/com.microsoft.Word/Data/runtime_logs.txt
```

> [!NOTE] 
> <span data-ttu-id="65a33-154">Vous devrez redémarrer Office après l’exécution de la commande `defaults` pour activer la journalisation de l’exécution.</span><span class="sxs-lookup"><span data-stu-id="65a33-154">You'll need to restart Office after running the `defaults` command to enable runtime logging.</span></span>

<span data-ttu-id="65a33-155">Pour désactiver la journalisation de l’exécution, utilisez la commande `defaults delete` :</span><span class="sxs-lookup"><span data-stu-id="65a33-155">To turn off runtime logging, use the `defaults delete` command:</span></span>

```command&nbsp;line
defaults delete <bundle id> CEFRuntimeLoggingFile
```

<span data-ttu-id="65a33-156">L’exemple suivant désactive la journalisation de l’exécution pour Word :</span><span class="sxs-lookup"><span data-stu-id="65a33-156">The following example will turn off runtime logging for Word:</span></span>

```command&nbsp;line
defaults delete com.microsoft.Word CEFRuntimeLoggingFile
```

### <a name="to-troubleshoot-issues-with-your-manifest"></a><span data-ttu-id="65a33-157">Résolution des problèmes avec votre manifeste</span><span class="sxs-lookup"><span data-stu-id="65a33-157">To troubleshoot issues with your manifest</span></span>

<span data-ttu-id="65a33-158">Pour utiliser la journalisation runtime pour résoudre les problèmes de chargement d’un complément, procédez comme suit :</span><span class="sxs-lookup"><span data-stu-id="65a33-158">To use runtime logging to troubleshoot issues loading an add-in:</span></span>
 
1. <span data-ttu-id="65a33-159">[Chargez une version test de votre complément](sideload-office-add-ins-for-testing.md).</span><span class="sxs-lookup"><span data-stu-id="65a33-159">[Sideload your add-in](sideload-office-add-ins-for-testing.md) for testing.</span></span> 

    > [!NOTE]
    > <span data-ttu-id="65a33-160">Nous vous recommandons de charger uniquement une version test du complément que vous testez pour réduire le nombre de messages dans le fichier journal.</span><span class="sxs-lookup"><span data-stu-id="65a33-160">We recommend that you sideload only the add-in that you are testing to minimize the number of messages in the log file.</span></span>

2. <span data-ttu-id="65a33-161">Si rien ne se produit et que votre complément n’apparaît pas (et ne s’affiche pas dans la boîte de dialogue des compléments), ouvrez le fichier journal.</span><span class="sxs-lookup"><span data-stu-id="65a33-161">If nothing happens and you don't see your add-in (and it's not appearing in the add-ins dialog box), open the log file.</span></span>

3. <span data-ttu-id="65a33-p112">Recherchez le fichier journal pour l’ID de votre complément, que vous définissez dans votre manifeste. Dans le fichier journal, cet ID est intitulé `SolutionId`.</span><span class="sxs-lookup"><span data-stu-id="65a33-p112">Search the log file for your add-in ID, which you define in your manifest. In the log file, this ID is labeled `SolutionId`.</span></span> 

<span data-ttu-id="65a33-p113">Dans l’exemple suivant, le fichier journal identifie un contrôle qui pointe vers un fichier de ressources qui n’existe pas. Pour cet exemple, la correction consistera à corriger la faute de frappe dans le manifeste ou à ajouter la ressource manquante.</span><span class="sxs-lookup"><span data-stu-id="65a33-p113">In the following example, the log file identifies a control that points to a resource file that doesn't exist. For this example, the fix would be to correct the typo in the manifest or to add the missing resource.</span></span>

![Capture d’écran d’un fichier journal avec une entrée qui spécifie un ID de ressource qui est introuvable](http://i.imgur.com/f8bouLA.png) 

### <a name="known-issues-with-runtime-logging"></a><span data-ttu-id="65a33-167">Problèmes connus avec la journalisation runtime</span><span class="sxs-lookup"><span data-stu-id="65a33-167">Known issues with runtime logging</span></span>

<span data-ttu-id="65a33-p114">Vous pouvez afficher des messages dans le fichier journal qui sont source de confusion ou classés de façon incorrecte. Par exemple :</span><span class="sxs-lookup"><span data-stu-id="65a33-p114">You might see messages in the log file that are confusing or that are classified incorrectly. For example:</span></span>

- <span data-ttu-id="65a33-170">Le message `Medium Current host not in add-in's host list` suivi de `Unexpected Parsed manifest targeting different host` est classé incorrectement en tant qu’erreur.</span><span class="sxs-lookup"><span data-stu-id="65a33-170">The message `Medium Current host not in add-in's host list` followed by `Unexpected Parsed manifest targeting different host` is incorrectly classified as an error.</span></span>

- <span data-ttu-id="65a33-171">Si vous voyez le message `Unexpected Add-in is missing required manifest fields DisplayName` et qu’il ne contient pas de SolutionId, l’erreur n’est probablement pas liée au complément que vous déboguez.</span><span class="sxs-lookup"><span data-stu-id="65a33-171">If you see the message `Unexpected Add-in is missing required manifest fields DisplayName` and it doesn't contain a SolutionId, the error is most likely not related to the add-in you are debugging.</span></span> 

- <span data-ttu-id="65a33-p115">Tous les messages `Monitorable` sont des erreurs attendues du point de vue du système. Parfois, ils indiquent un problème avec votre manifeste, comme un élément mal orthographié qui a été ignoré, mais n’a pas provoqué l’échec du manifeste.</span><span class="sxs-lookup"><span data-stu-id="65a33-p115">Any `Monitorable` messages are expected errors from a system point of view. Sometimes they indicate an issue with your manifest, such as a misspelled element that was skipped but didn't cause the manifest to fail.</span></span> 

## <a name="clear-the-office-cache"></a><span data-ttu-id="65a33-174">Vider le cache Office</span><span class="sxs-lookup"><span data-stu-id="65a33-174">Clear the Office cache</span></span>

<span data-ttu-id="65a33-175">Si les modifications apportées au manifeste, par exemple aux noms de fichier des icônes de bouton dans le ruban ou au texte des commandes de complément, ne semblent pas être appliquées, essayez de vider le cache Office de votre ordinateur.</span><span class="sxs-lookup"><span data-stu-id="65a33-175">If changes you've made in the manifest, such as file names of ribbon button icons, or text of add-in commands, do not seem to take effect, try clearing the Office cache on your computer.</span></span> 

#### <a name="for-windows"></a><span data-ttu-id="65a33-176">Pour Windows :</span><span class="sxs-lookup"><span data-stu-id="65a33-176">For Windows:</span></span>
<span data-ttu-id="65a33-177">Supprimer le contenu du dossier `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\`.</span><span class="sxs-lookup"><span data-stu-id="65a33-177">Delete the content of the folder `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\`.</span></span>

#### <a name="for-mac"></a><span data-ttu-id="65a33-178">Pour Mac :</span><span class="sxs-lookup"><span data-stu-id="65a33-178">For Mac:</span></span>

[!include[additional cache folders on Mac](../includes/mac-cache-folders.md)]

#### <a name="for-ios"></a><span data-ttu-id="65a33-179">Pour iOS :</span><span class="sxs-lookup"><span data-stu-id="65a33-179">For iOS:</span></span>
<span data-ttu-id="65a33-p116">Appelez `window.location.reload(true)` à partir de JavaScript dans le complément pour forcer le rechargement. Vous pouvez également choisir de réinstaller Office.</span><span class="sxs-lookup"><span data-stu-id="65a33-p116">Call `window.location.reload(true)` from JavaScript in the add-in to force a reload. Alternatively, you can reinstall Office.</span></span>

## <a name="see-also"></a><span data-ttu-id="65a33-182">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="65a33-182">See also</span></span>

- [<span data-ttu-id="65a33-183">Manifeste XML des compléments Office</span><span class="sxs-lookup"><span data-stu-id="65a33-183">Office Add-ins XML manifest</span></span>](../develop/add-in-manifests.md)
- [<span data-ttu-id="65a33-184">Chargement de la version test des compléments Office</span><span class="sxs-lookup"><span data-stu-id="65a33-184">Sideload Office Add-ins for testing</span></span>](sideload-office-add-ins-for-testing.md)
- [<span data-ttu-id="65a33-185">Débogage des compléments Office</span><span class="sxs-lookup"><span data-stu-id="65a33-185">Debug Office Add-ins</span></span>](debug-add-ins-using-f12-developer-tools-on-windows-10.md)
