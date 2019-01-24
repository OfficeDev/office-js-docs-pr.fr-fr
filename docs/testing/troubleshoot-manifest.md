---
title: Valider et résoudre des problèmes avec votre manifeste
description: Utiliser ces méthodes pour valider le manifeste des compléments Office.
ms.date: 11/02/2018
localization_priority: Priority
ms.openlocfilehash: 921adf6f1f398887d96031790facc1fb1425af2b
ms.sourcegitcommit: d1aa7201820176ed986b9f00bb9c88e055906c77
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 01/23/2019
ms.locfileid: "29388374"
---
# <a name="validate-and-troubleshoot-issues-with-your-manifest"></a><span data-ttu-id="26abf-103">Valider et résoudre des problèmes avec votre manifeste</span><span class="sxs-lookup"><span data-stu-id="26abf-103">Validate and troubleshoot issues with your manifest</span></span>

<span data-ttu-id="26abf-104">Utiliser les méthodes suivantes pour valider et résoudre les problèmes rencontrés dans votre manifeste pour compléments Office :</span><span class="sxs-lookup"><span data-stu-id="26abf-104">Use these methods to validate and troubleshoot issues in your Office Add-ins manifest:</span></span> 

- [<span data-ttu-id="26abf-105">Validation du manifeste à l’aide du validateur de complément Office</span><span class="sxs-lookup"><span data-stu-id="26abf-105">Validate your manifest with the Office Add-in Validator</span></span>](#validate-your-manifest-with-the-office-add-in-validator)   
- [<span data-ttu-id="26abf-106">Validation de votre manifeste par rapport au schéma XML</span><span class="sxs-lookup"><span data-stu-id="26abf-106">Validate your manifest against the XML schema</span></span>](#validate-your-manifest-against-the-xml-schema)
- [<span data-ttu-id="26abf-107">Validation de votre manifeste avec le générateur Yeoman pour les compléments Office</span><span class="sxs-lookup"><span data-stu-id="26abf-107">Validate your manifest with the Yeoman generator for Office Add-ins</span></span>](#validate-your-manifest-with-the-yeoman-generator-for-office-add-ins)
- [<span data-ttu-id="26abf-108">Utilisation de la journalisation runtime pour déboguer votre complément</span><span class="sxs-lookup"><span data-stu-id="26abf-108">Use runtime logging to debug your add-in</span></span>](#use-runtime-logging-to-debug-your-add-in)


## <a name="validate-your-manifest-with-the-office-add-in-validator"></a><span data-ttu-id="26abf-109">Validation du manifeste à l’aide du validateur de complément Office</span><span class="sxs-lookup"><span data-stu-id="26abf-109">Validate your manifest with the Office Add-in Validator</span></span>

<span data-ttu-id="26abf-110">Pour vous assurer que le fichier manifeste qui décrit votre complément Office est correct et complet, vérifiez-le à l’aide du [validateur de complément Office](https://github.com/OfficeDev/office-addin-validator).</span><span class="sxs-lookup"><span data-stu-id="26abf-110">To help ensure that the manifest file that describes your Office Add-in is correct and complete, validate it against the [Office Add-in Validator](https://github.com/OfficeDev/office-addin-validator).</span></span>

### <a name="to-use-the-office-add-in-validator-to-validate-your-manifest"></a><span data-ttu-id="26abf-111">Pour utiliser le validateur de complément Office afin de valider votre manifeste</span><span class="sxs-lookup"><span data-stu-id="26abf-111">To use the Office Add-in Validator to validate your manifest</span></span>

1. <span data-ttu-id="26abf-112">Installez [Node.js](https://nodejs.org/download/).</span><span class="sxs-lookup"><span data-stu-id="26abf-112">Install [Node.js](https://nodejs.org/download/).</span></span> 

2. <span data-ttu-id="26abf-113">Ouvrez une invite de commandes/un terminal en tant qu’administrateur, puis installez le validateur de complément Office et ses dépendances de façon globale à l’aide de la commande suivante :</span><span class="sxs-lookup"><span data-stu-id="26abf-113">Open a command prompt / terminal as an administrator, and install the Office Add-in Validator and its dependencies globally by using the following command:</span></span>

    ```bash
    npm install -g office-addin-validator
    ```
    
    > [!NOTE]
    > <span data-ttu-id="26abf-114">Si Yo Office est déjà installé, effectuez une mise à niveau vers la dernière version ; le validateur sera installé en tant que dépendance.</span><span class="sxs-lookup"><span data-stu-id="26abf-114">If you already have Yo Office installed, upgrade to the latest version, and the validator will be installed as a dependency.</span></span>

3. <span data-ttu-id="26abf-p101">Exécutez la commande suivante pour valider votre manifeste. Remplacez MANIFEST.XML par le chemin d’accès au fichier XML de manifeste.</span><span class="sxs-lookup"><span data-stu-id="26abf-p101">Run the following command to validate your manifest. Replace MANIFEST.XML with the path to the manifest XML file.</span></span>

    ```bash
    validate-office-addin MANIFEST.XML
    ```

## <a name="validate-your-manifest-against-the-xml-schema"></a><span data-ttu-id="26abf-117">Validation de votre manifeste par rapport au schéma XML</span><span class="sxs-lookup"><span data-stu-id="26abf-117">Validate your manifest against the XML schema</span></span>

<span data-ttu-id="26abf-118">Assurez-vous que le fichier manifeste suit le schéma approprié, y compris les espaces de noms pour les éléments que vous utilisez.</span><span class="sxs-lookup"><span data-stu-id="26abf-118">To help ensure that the manifest file follows the correct schema, including any namespaces for the elements you are using.</span></span> <span data-ttu-id="26abf-119">Si vous avez copié des éléments à partir d’autres exemples de manifestes, vérifiez par deux fois que vous avez également **inclus les espaces de noms appropriés**.</span><span class="sxs-lookup"><span data-stu-id="26abf-119">If you copied elements from other sample manifests double check you also **include the appropriate namespaces**.</span></span> <span data-ttu-id="26abf-120">Vous pouvez valider un manifeste par rapport aux fichiers de [définition de schéma XML (XSD)](https://github.com/OfficeDev/office-js-docs-pr/tree/master/docs/overview/schemas).</span><span class="sxs-lookup"><span data-stu-id="26abf-120">You can validate a manifest against the [XML Schema Definition (XSD)](https://github.com/OfficeDev/office-js-docs-pr/tree/master/docs/overview/schemas) files.</span></span> <span data-ttu-id="26abf-121">Pour ce faire, vous pouvez utiliser un outil de validation de schéma XML.</span><span class="sxs-lookup"><span data-stu-id="26abf-121">You can use an XML schema validation tool to perform this validation.</span></span> 



### <a name="to-use-a-command-line-xml-schema-validation-tool-to-validate-your-manifest"></a><span data-ttu-id="26abf-122">Pour utiliser un outil de validation de schéma XML à ligne de commande pour valider votre manifeste</span><span class="sxs-lookup"><span data-stu-id="26abf-122">To use a command-line XML schema validation tool to validate your manifest</span></span>

1.  <span data-ttu-id="26abf-123">Installez [tar](https://www.gnu.org/software/tar/) et [libxml](http://xmlsoft.org/FAQ.html), si vous ne l’avez pas déjà fait.</span><span class="sxs-lookup"><span data-stu-id="26abf-123">Install [tar](https://www.gnu.org/software/tar/) and [libxml](http://xmlsoft.org/FAQ.html), if you haven't already.</span></span>

2.  <span data-ttu-id="26abf-p103">Exécutez la commande suivante. Remplacez `XSD_FILE` par le chemin d’accès au fichier XSD manifeste et `XML_FILE` par le chemin d’accès au fichier XML manifeste.</span><span class="sxs-lookup"><span data-stu-id="26abf-p103">Run the following command. Replace `XSD_FILE` with the path to the manifest XSD file, and replace `XML_FILE` with the path to the manifest XML file.</span></span>
    
    ```bash
    xmllint --noout --schema XSD_FILE XML_FILE
    ```

## <a name="validate-your-manifest-with-the-yeoman-generator-for-office-add-ins"></a><span data-ttu-id="26abf-126">Validation de votre manifeste avec le générateur Yeoman pour les compléments Office</span><span class="sxs-lookup"><span data-stu-id="26abf-126">Validate your manifest with the Yeoman generator for Office Add-ins</span></span>

<span data-ttu-id="26abf-127">Si vous avez créé votre complément Office à l’aide du [générateur Yeoman pour les compléments Office](https://www.npmjs.com/package/generator-office), vérifiez que le fichier manifeste suit le schéma correct en exécutant la commande suivante dans le répertoire racine de votre projet :</span><span class="sxs-lookup"><span data-stu-id="26abf-127">If you've created your Office Add-in using the [Yeoman generator for Office Add-ins](https://www.npmjs.com/package/generator-office), you can ensure that the manifest file follows the correct schema by running the following command within the root directory of your project:</span></span>

```bash
npm run validate
```

![Gif animé qui montre le validateur Yo Office exécuté sur la ligne de commande et les résultats générés indiquant « Validation Passed » (validation réussie)](../images/yo-office-validator.gif)

> [!NOTE]
> <span data-ttu-id="26abf-129">Pour accéder à cette fonctionnalité, votre projet de complément doit être créé à l’aide du [générateur Yeoman pour les compléments Office](https://www.npmjs.com/package/generator-office) (version 1.1.17 ou ultérieure).</span><span class="sxs-lookup"><span data-stu-id="26abf-129">To have access to this functionality, your add-in project must have been created by using [Yeoman generator for Office Add-ins](https://www.npmjs.com/package/generator-office) version 1.1.17 or later.</span></span>

## <a name="use-runtime-logging-to-debug-your-add-in"></a><span data-ttu-id="26abf-130">Utilisation de la journalisation runtime pour déboguer votre complément</span><span class="sxs-lookup"><span data-stu-id="26abf-130">Use runtime logging to debug your add-in</span></span> 

<span data-ttu-id="26abf-131">Vous pouvez utiliser la journalisation runtime pour déboguer le manifeste de votre complément ainsi que plusieurs erreurs d’installation.</span><span class="sxs-lookup"><span data-stu-id="26abf-131">You can use runtime logging to debug your add-in's manifest as well as several installation errors.</span></span> <span data-ttu-id="26abf-132">Cette fonctionnalité peut vous aider à identifier et à résoudre les problèmes avec votre manifeste qui ne sont pas détectés par la validation de schéma XSD, comme une incompatibilité entre les ID de ressources.</span><span class="sxs-lookup"><span data-stu-id="26abf-132">This feature can help you identify and fix issues with your manifest that are not detected by XSD schema validation, such as a mismatch between resource IDs.</span></span> <span data-ttu-id="26abf-133">La journalisation runtime est particulièrement utile pour déboguer des compléments qui implémentent des commandes de complément et des fonctions personnalisées Excel.</span><span class="sxs-lookup"><span data-stu-id="26abf-133">Runtime logging is particularly  useful for debugging add-ins that implement add-in commands and Excel custom functions.</span></span>   

> [!NOTE]
> <span data-ttu-id="26abf-134">La fonctionnalité de journalisation runtime est actuellement disponible pour Office 2016 pour ordinateur de bureau.</span><span class="sxs-lookup"><span data-stu-id="26abf-134">The runtime logging feature is currently available for Office 2016 desktop.</span></span>

### <a name="to-turn-on-runtime-logging"></a><span data-ttu-id="26abf-135">Pour activer la journalisation runtime</span><span class="sxs-lookup"><span data-stu-id="26abf-135">To turn on runtime logging</span></span>

> [!IMPORTANT]
> <span data-ttu-id="26abf-p105">La journalisation runtime réduit les performances. Activez-la uniquement lorsque vous avez besoin de déboguer des problèmes avec votre manifeste de complément.</span><span class="sxs-lookup"><span data-stu-id="26abf-p105">Runtime Logging affects performance. Turn it on only when you need to debug issues with your add-in manifest.</span></span>

<span data-ttu-id="26abf-138">Pour activer la journalisation runtime, procédez comme suit :</span><span class="sxs-lookup"><span data-stu-id="26abf-138">To turn on runtime logging:</span></span>

1. <span data-ttu-id="26abf-139">Vérifiez que vous exécutez la version Bureau d’Office 2016 **16.0.7019** ou une version ultérieure.</span><span class="sxs-lookup"><span data-stu-id="26abf-139">Make sure that you are running Office 2016 desktop build **16.0.7019** or later.</span></span> 

2. <span data-ttu-id="26abf-140">Ajoutez la clé de registre `RuntimeLogging` sous `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\WEF\Developer\`.</span><span class="sxs-lookup"><span data-stu-id="26abf-140">Add the `RuntimeLogging` registry key under `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\WEF\Developer\`.</span></span> 

    > [!NOTE]
    > <span data-ttu-id="26abf-141">Si la clé (dossier) `Developer` n’existe pas sous `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\WEF\`, procédez comme suit pour la créer :</span><span class="sxs-lookup"><span data-stu-id="26abf-141">If the `Developer` key (folder) does not already exist under `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\WEF\`, complete the following steps to create it:</span></span> 
    > 1. <span data-ttu-id="26abf-142">Cliquez avec le bouton droit de votre souris sur la clé (dossier) **WEF**, puis sélectionnez **Nouveau** > **Clé**.</span><span class="sxs-lookup"><span data-stu-id="26abf-142">Right-click the **WEF** key (folder) and select **New** > **Key**.</span></span>
    > 2. <span data-ttu-id="26abf-143">Nommez la nouvelle clé **Développeur**.</span><span class="sxs-lookup"><span data-stu-id="26abf-143">Name the new key **Developer**.</span></span>

3. <span data-ttu-id="26abf-p106">Définissez la valeur par défaut de la clé pour le chemin d’accès complet du fichier dans lequel écrire le journal. Pour obtenir un exemple, reportez-vous à [EnableRuntimeLogging.zip](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/raw/master/Tools/RuntimeLogging/EnableRuntimeLogging.zip).</span><span class="sxs-lookup"><span data-stu-id="26abf-p106">Set the default value of the key to the full path of the file where you want the log to be written. For an example, see [EnableRuntimeLogging.zip](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/raw/master/Tools/RuntimeLogging/EnableRuntimeLogging.zip).</span></span> 

    > [!NOTE]
    > <span data-ttu-id="26abf-146">Le répertoire dans lequel le fichier journal sera écrit doit déjà exister et vous devez disposer des autorisations d’écriture correspondantes.</span><span class="sxs-lookup"><span data-stu-id="26abf-146">The directory in which the log file will be written must already exist, and you must have write permissions to it.</span></span> 
 
<span data-ttu-id="26abf-147">L’image suivante indique à quoi doit ressembler le registre.</span><span class="sxs-lookup"><span data-stu-id="26abf-147">The following image shows what the registry should look like.</span></span> <span data-ttu-id="26abf-148">Pour désactiver la fonctionnalité, supprimez la clé de registre `RuntimeLogging`.</span><span class="sxs-lookup"><span data-stu-id="26abf-148">To turn the feature off, remove the `RuntimeLogging` key from the registry.</span></span> 

![Capture d’écran de l’Éditeur du Registre avec la clé de registre RuntimeLogging](http://i.imgur.com/Sa9TyI6.png)


### <a name="to-troubleshoot-issues-with-your-manifest"></a><span data-ttu-id="26abf-150">Résolution des problèmes avec votre manifeste</span><span class="sxs-lookup"><span data-stu-id="26abf-150">To troubleshoot issues with your manifest</span></span>

<span data-ttu-id="26abf-151">Pour utiliser la journalisation runtime pour résoudre les problèmes de chargement d’un complément, procédez comme suit :</span><span class="sxs-lookup"><span data-stu-id="26abf-151">To use runtime logging to troubleshoot issues loading an add-in:</span></span>
 
1. <span data-ttu-id="26abf-152">[Chargez une version test de votre complément](sideload-office-add-ins-for-testing.md).</span><span class="sxs-lookup"><span data-stu-id="26abf-152">[Sideload your add-in](sideload-office-add-ins-for-testing.md) for testing.</span></span> 

    > [!NOTE]
    > <span data-ttu-id="26abf-153">Nous vous recommandons de charger uniquement une version test du complément que vous testez pour réduire le nombre de messages dans le fichier journal.</span><span class="sxs-lookup"><span data-stu-id="26abf-153">We recommend that you sideload only the add-in that you are testing to minimize the number of messages in the log file.</span></span>

2. <span data-ttu-id="26abf-154">Si rien ne se produit et que votre complément n’apparaît pas (et ne s’affiche pas dans la boîte de dialogue des compléments), ouvrez le fichier journal.</span><span class="sxs-lookup"><span data-stu-id="26abf-154">If nothing happens and you don't see your add-in (and it's not appearing in the add-ins dialog box), open the log file.</span></span>

3. <span data-ttu-id="26abf-p108">Recherchez le fichier journal pour l’ID de votre complément, que vous définissez dans votre manifeste. Dans le fichier journal, cet ID est intitulé `SolutionId`.</span><span class="sxs-lookup"><span data-stu-id="26abf-p108">Search the log file for your add-in ID, which you define in your manifest. In the log file, this ID is labeled `SolutionId`.</span></span> 

<span data-ttu-id="26abf-p109">Dans l’exemple suivant, le fichier journal identifie un contrôle qui pointe vers un fichier de ressources qui n’existe pas. Pour cet exemple, la correction consistera à corriger la faute de frappe dans le manifeste ou à ajouter la ressource manquante.</span><span class="sxs-lookup"><span data-stu-id="26abf-p109">In the following example, the log file identifies a control that points to a resource file that doesn't exist. For this example, the fix would be to correct the typo in the manifest or to add the missing resource.</span></span>

![Capture d’écran d’un fichier journal avec une entrée qui spécifie un ID de ressource qui est introuvable](http://i.imgur.com/f8bouLA.png) 

### <a name="known-issues-with-runtime-logging"></a><span data-ttu-id="26abf-160">Problèmes connus avec la journalisation runtime</span><span class="sxs-lookup"><span data-stu-id="26abf-160">Known issues with runtime logging</span></span>

<span data-ttu-id="26abf-p110">Vous pouvez afficher des messages dans le fichier journal qui sont source de confusion ou classés de façon incorrecte. Par exemple :</span><span class="sxs-lookup"><span data-stu-id="26abf-p110">You might see messages in the log file that are confusing or that are classified incorrectly. For example:</span></span>

- <span data-ttu-id="26abf-163">Le message `Medium Current host not in add-in's host list` suivi de `Unexpected Parsed manifest targeting different host` est classé incorrectement en tant qu’erreur.</span><span class="sxs-lookup"><span data-stu-id="26abf-163">The message `Medium Current host not in add-in's host list` followed by `Unexpected Parsed manifest targeting different host` is incorrectly classified as an error.</span></span>

- <span data-ttu-id="26abf-164">Si vous voyez le message `Unexpected Add-in is missing required manifest fields DisplayName` et qu’il ne contient pas de SolutionId, l’erreur n’est probablement pas liée au complément que vous déboguez.</span><span class="sxs-lookup"><span data-stu-id="26abf-164">If you see the message `Unexpected Add-in is missing required manifest fields DisplayName` and it doesn't contain a SolutionId, the error is most likely not related to the add-in you are debugging.</span></span> 

- <span data-ttu-id="26abf-p111">Tous les messages `Monitorable` sont des erreurs attendues du point de vue du système. Parfois, ils indiquent un problème avec votre manifeste, comme un élément mal orthographié qui a été ignoré, mais n’a pas provoqué l’échec du manifeste.</span><span class="sxs-lookup"><span data-stu-id="26abf-p111">Any `Monitorable` messages are expected errors from a system point of view. Sometimes they indicate an issue with your manifest, such as a misspelled element that was skipped but didn't cause the manifest to fail.</span></span> 

## <a name="clear-the-office-cache"></a><span data-ttu-id="26abf-167">Vider le cache Office</span><span class="sxs-lookup"><span data-stu-id="26abf-167">Clear the Office cache</span></span>

<span data-ttu-id="26abf-168">Si les modifications apportées au manifeste, par exemple aux noms de fichier des icônes de bouton dans le ruban ou au texte des commandes de complément, ne semblent pas appliquées, essayez de vider le cache Office de votre ordinateur.</span><span class="sxs-lookup"><span data-stu-id="26abf-168">If changes you've made in the manifest, such as file names of ribbon button icons, or text of add-in commands, do not seem to take effect, try clearing the Office cache on your computer.</span></span> 

#### <a name="for-windows"></a><span data-ttu-id="26abf-169">Pour Windows :</span><span class="sxs-lookup"><span data-stu-id="26abf-169">For Windows:</span></span>
<span data-ttu-id="26abf-170">Supprimez le contenu du dossier `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\`.</span><span class="sxs-lookup"><span data-stu-id="26abf-170">Delete the content of the folder `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\`.</span></span>

#### <a name="for-mac"></a><span data-ttu-id="26abf-171">Pour Mac :</span><span class="sxs-lookup"><span data-stu-id="26abf-171">For Mac:</span></span>
<span data-ttu-id="26abf-172">Supprimez le contenu du dossier `/Users/{your_name_on_the_device}/Library/Containers/com.Microsoft.OsfWebHost/Data/`.</span><span class="sxs-lookup"><span data-stu-id="26abf-172">Delete the content of the folder `/Users/{your_name_on_the_device}/Library/Containers/com.Microsoft.OsfWebHost/Data/`.</span></span>

#### <a name="for-ios"></a><span data-ttu-id="26abf-173">Pour iOS :</span><span class="sxs-lookup"><span data-stu-id="26abf-173">For iOS:</span></span>
<span data-ttu-id="26abf-p112">Appelez `window.location.reload(true)` à partir de JavaScript dans le complément pour forcer le rechargement. Vous pouvez également choisir de réinstaller Office.</span><span class="sxs-lookup"><span data-stu-id="26abf-p112">Call `window.location.reload(true)` from JavaScript in the add-in to force a reload. Alternatively, you can reinstall Office.</span></span>

## <a name="see-also"></a><span data-ttu-id="26abf-176">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="26abf-176">See also</span></span>

- [<span data-ttu-id="26abf-177">Manifeste XML des compléments Office</span><span class="sxs-lookup"><span data-stu-id="26abf-177">Office Add-ins XML manifest</span></span>](../develop/add-in-manifests.md)
- [<span data-ttu-id="26abf-178">Chargement de la version test des compléments Office</span><span class="sxs-lookup"><span data-stu-id="26abf-178">Sideload Office Add-ins for testing</span></span>](sideload-office-add-ins-for-testing.md)
- [<span data-ttu-id="26abf-179">Débogage des compléments Office</span><span class="sxs-lookup"><span data-stu-id="26abf-179">Debug Office Add-ins</span></span>](debug-add-ins-using-f12-developer-tools-on-windows-10.md)
