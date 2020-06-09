---
title: Déboguez votre complément avec la journalisation runtime
description: Découvrez l’utilisation de la journalisation runtime pour déboguer votre complément.
ms.date: 12/31/2019
localization_priority: Normal
ms.openlocfilehash: e97f989ff4791da678fbecbf8bceed3670749e01
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/08/2020
ms.locfileid: "44611146"
---
# <a name="debug-your-add-in-with-runtime-logging"></a><span data-ttu-id="6168a-103">Déboguez votre complément avec la journalisation runtime</span><span class="sxs-lookup"><span data-stu-id="6168a-103">Debug your add-in with runtime logging</span></span>

<span data-ttu-id="6168a-104">Vous pouvez utiliser la journalisation runtime pour déboguer le manifeste de votre complément ainsi que plusieurs erreurs d’installation.</span><span class="sxs-lookup"><span data-stu-id="6168a-104">You can use runtime logging to debug your add-in's manifest as well as several installation errors.</span></span> <span data-ttu-id="6168a-105">Cette fonctionnalité peut vous aider à identifier et à résoudre les problèmes avec votre manifeste qui ne sont pas détectés par la validation de schéma XSD, comme une incompatibilité entre les ID de ressources.</span><span class="sxs-lookup"><span data-stu-id="6168a-105">This feature can help you identify and fix issues with your manifest that are not detected by XSD schema validation, such as a mismatch between resource IDs.</span></span> <span data-ttu-id="6168a-106">La journalisation runtime est particulièrement utile pour déboguer des compléments qui implémentent des commandes de complément et des fonctions personnalisées Excel.</span><span class="sxs-lookup"><span data-stu-id="6168a-106">Runtime logging is particularly  useful for debugging add-ins that implement add-in commands and Excel custom functions.</span></span>   

> [!NOTE]
> <span data-ttu-id="6168a-107">La fonctionnalité de journalisation runtime est actuellement disponible pour Office 2016 pour ordinateur de bureau.</span><span class="sxs-lookup"><span data-stu-id="6168a-107">The runtime logging feature is currently available for Office 2016 desktop.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="6168a-108">La journalisation runtime affecte les performances.</span><span class="sxs-lookup"><span data-stu-id="6168a-108">Runtime Logging affects performance.</span></span> <span data-ttu-id="6168a-109">Activez-la uniquement lorsque vous avez besoin de déboguer des problèmes avec votre manifeste de complément.</span><span class="sxs-lookup"><span data-stu-id="6168a-109">Turn it on only when you need to debug issues with your add-in manifest.</span></span>

## <a name="use-runtime-logging-from-the-command-line"></a><span data-ttu-id="6168a-110">Utiliser la journalisation de l’exécution à partir de la ligne de commande</span><span class="sxs-lookup"><span data-stu-id="6168a-110">Use runtime logging from the command line</span></span>

<span data-ttu-id="6168a-111">L’activation de la journalisation de l’exécution à partir de la ligne de commande est le moyen le plus rapide d’utiliser cet outil de journalisation.</span><span class="sxs-lookup"><span data-stu-id="6168a-111">Enabling runtime logging from the command line is the fastest way to use this logging tool.</span></span> <span data-ttu-id="6168a-112">Celles-ci utilisent npx, fourni par défaut dans le cadre de npm@5.2.0 +.</span><span class="sxs-lookup"><span data-stu-id="6168a-112">These use npx, which is provided by default as part of npm@5.2.0+.</span></span> <span data-ttu-id="6168a-113">Si vous disposez d’une version antérieure de [npm](https://www.npmjs.com/), essayez les instructions [Journalisation de l’exécution sur Windows](#runtime-logging-on-windows) ou [Journalisation de l’exécution sur Mac](#runtime-logging-on-mac), ou [install npx](https://www.npmjs.com/package/npx).</span><span class="sxs-lookup"><span data-stu-id="6168a-113">If you have an earlier version of [npm](https://www.npmjs.com/), try [Runtime logging on Windows](#runtime-logging-on-windows) or [Runtime logging on Mac](#runtime-logging-on-mac) instructions, or [install npx](https://www.npmjs.com/package/npx).</span></span>

- <span data-ttu-id="6168a-114">Pour activer la journalisation de l’exécution :</span><span class="sxs-lookup"><span data-stu-id="6168a-114">To enable runtime logging:</span></span>
    ```command&nbsp;line
    npx office-addin-dev-settings runtime-log --enable
    ```
- <span data-ttu-id="6168a-115">Pour activer la journalisation de l’exécution uniquement pour un fichier spécifique, utilisez la même commande avec un nom de fichier :</span><span class="sxs-lookup"><span data-stu-id="6168a-115">To enable runtime logging only for a specific file, use the same command with a filename:</span></span>

    ```command&nbsp;line
    npx office-addin-dev-settings runtime-log --enable [filename.txt]
    ```

- <span data-ttu-id="6168a-116">Pour désactiver la journalisation de l’exécution :</span><span class="sxs-lookup"><span data-stu-id="6168a-116">To disable runtime logging:</span></span>

    ```command&nbsp;line
    npx office-addin-dev-settings runtime-log --disable
    ```

- <span data-ttu-id="6168a-117">Pour indiquer si la journalisation de l’exécution est activée :</span><span class="sxs-lookup"><span data-stu-id="6168a-117">To display whether runtime logging is enabled:</span></span>

    ```command&nbsp;line
    npx office-addin-dev-settings runtime-log
    ```

- <span data-ttu-id="6168a-118">Pour afficher l’aide au sein de la ligne de commande pour la journalisation de l’exécution :</span><span class="sxs-lookup"><span data-stu-id="6168a-118">To display help within the command line for runtime logging:</span></span>

    ```command&nbsp;line
    npx office-addin-dev-settings runtime-log --help
    ```

## <a name="runtime-logging-on-windows"></a><span data-ttu-id="6168a-119">Journalisation de l’exécution sur Windows</span><span class="sxs-lookup"><span data-stu-id="6168a-119">Runtime logging on Windows</span></span>

1. <span data-ttu-id="6168a-120">Vérifiez que vous exécutez la version Bureau d’Office 2016 **16.0.7019** ou une version ultérieure.</span><span class="sxs-lookup"><span data-stu-id="6168a-120">Make sure that you are running Office 2016 desktop build **16.0.7019** or later.</span></span> 

2. <span data-ttu-id="6168a-121">Ajoutez la clé de registre `RuntimeLogging` sous `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\WEF\Developer\`.</span><span class="sxs-lookup"><span data-stu-id="6168a-121">Add the `RuntimeLogging` registry key under `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\WEF\Developer\`.</span></span> 

    > [!NOTE]
    > <span data-ttu-id="6168a-122">Si la clé (dossier) `Developer` n’existe pas sous `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\WEF\`, procédez comme suit pour la créer :</span><span class="sxs-lookup"><span data-stu-id="6168a-122">If the `Developer` key (folder) does not already exist under `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\WEF\`, complete the following steps to create it:</span></span> 
    > 1. <span data-ttu-id="6168a-123">Cliquez avec le bouton droit de votre souris sur la clé (dossier) **WEF**, puis sélectionnez **Nouveau** > **Clé**.</span><span class="sxs-lookup"><span data-stu-id="6168a-123">Right-click the **WEF** key (folder) and select **New** > **Key**.</span></span>
    > 2. <span data-ttu-id="6168a-124">Nommez la nouvelle clé **Développeur**.</span><span class="sxs-lookup"><span data-stu-id="6168a-124">Name the new key **Developer**.</span></span>

3. <span data-ttu-id="6168a-125">Définissez la valeur par défaut de la clé **RuntimeLogging** pour le chemin d’accès complet du fichier dans lequel écrire le journal.</span><span class="sxs-lookup"><span data-stu-id="6168a-125">Set the default value of the **RuntimeLogging** key to the full path of the file where you want the log to be written.</span></span> <span data-ttu-id="6168a-126">Pour obtenir un exemple, voir [EnableRuntimeLogging.zip](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/raw/master/Tools/RuntimeLogging/EnableRuntimeLogging.zip).</span><span class="sxs-lookup"><span data-stu-id="6168a-126">For an example, see [EnableRuntimeLogging.zip](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/raw/master/Tools/RuntimeLogging/EnableRuntimeLogging.zip).</span></span> 

    > [!NOTE]
    > <span data-ttu-id="6168a-127">Le répertoire dans lequel le fichier journal sera écrit doit déjà exister et vous devez disposer des autorisations d’écriture correspondantes.</span><span class="sxs-lookup"><span data-stu-id="6168a-127">The directory in which the log file will be written must already exist, and you must have write permissions to it.</span></span> 
 
<span data-ttu-id="6168a-p105">L’image suivante indique à quoi doit ressembler le registre. Pour désactiver la fonctionnalité, supprimez la clé de registre `RuntimeLogging`.</span><span class="sxs-lookup"><span data-stu-id="6168a-p105">The following image shows what the registry should look like. To turn the feature off, remove the `RuntimeLogging` key from the registry.</span></span> 

![Capture d’écran de l’Éditeur du Registre avec la clé de registre RuntimeLogging](http://i.imgur.com/Sa9TyI6.png)

## <a name="runtime-logging-on-mac"></a><span data-ttu-id="6168a-131">Journalisation de l’exécution sur Mac</span><span class="sxs-lookup"><span data-stu-id="6168a-131">Runtime logging on Mac</span></span>

1. <span data-ttu-id="6168a-132">Vérifiez que vous exécutez la version Bureau d’Office 2016 **16.27** (19071500) ou une version ultérieure.</span><span class="sxs-lookup"><span data-stu-id="6168a-132">Make sure that you are running Office 2016 desktop build **16.27** (19071500) or later.</span></span>

2. <span data-ttu-id="6168a-133">Ouvrez **Terminal** et configurez une préférence de journalisation de l’exécution à l’aide de la commande `defaults` :</span><span class="sxs-lookup"><span data-stu-id="6168a-133">Open **Terminal** and set a runtime logging preference by using the `defaults` command:</span></span>
    
    ```command&nbsp;line
    defaults write <bundle id> CEFRuntimeLoggingFile -string <file_name>
    ```

    <span data-ttu-id="6168a-134">`<bundle id>` identifie l’hôte pour lequel activer la journalisation de l’exécution.</span><span class="sxs-lookup"><span data-stu-id="6168a-134">`<bundle id>` identifies which the host for which to enable runtime logging.</span></span> <span data-ttu-id="6168a-135">`<file_name>` est le nom du fichier texte dans lequel le journal sera écrit.</span><span class="sxs-lookup"><span data-stu-id="6168a-135">`<file_name>` is the name of the text file to which the log will be written.</span></span>

    <span data-ttu-id="6168a-136">Configurez `<bundle id>` à l’une des valeurs suivantes pour activer la journalisation de l’exécution pour l’hôte correspondant :</span><span class="sxs-lookup"><span data-stu-id="6168a-136">Set `<bundle id>` to one of the following values to enable runtime logging for the corresponding host:</span></span>

    - `com.microsoft.Word`
    - `com.microsoft.Excel`
    - `com.microsoft.Powerpoint`
    - `com.microsoft.Outlook`

<span data-ttu-id="6168a-137">L’exemple suivant montre comment activer la journalisation de l’exécution pour Word, puis ouvrir le fichier journal :</span><span class="sxs-lookup"><span data-stu-id="6168a-137">The following example enables runtime logging for Word and then opens the log file:</span></span>

```command&nbsp;line
defaults write com.microsoft.Word CEFRuntimeLoggingFile -string "runtime_logs.txt"
open ~/library/Containers/com.microsoft.Word/Data/runtime_logs.txt
```

> [!NOTE] 
> <span data-ttu-id="6168a-138">Vous devrez redémarrer Office après l’exécution de la commande `defaults` pour activer la journalisation de l’exécution.</span><span class="sxs-lookup"><span data-stu-id="6168a-138">You'll need to restart Office after running the `defaults` command to enable runtime logging.</span></span>

<span data-ttu-id="6168a-139">Pour désactiver la journalisation de l’exécution, utilisez la commande `defaults delete` :</span><span class="sxs-lookup"><span data-stu-id="6168a-139">To turn off runtime logging, use the `defaults delete` command:</span></span>

```command&nbsp;line
defaults delete <bundle id> CEFRuntimeLoggingFile
```

<span data-ttu-id="6168a-140">L’exemple suivant désactive la journalisation de l’exécution pour Word :</span><span class="sxs-lookup"><span data-stu-id="6168a-140">The following example will turn off runtime logging for Word:</span></span>

```command&nbsp;line
defaults delete com.microsoft.Word CEFRuntimeLoggingFile
```

## <a name="use-runtime-logging-to-troubleshoot-issues-with-your-manifest"></a><span data-ttu-id="6168a-141">Utilisez la journalisation runtime pour résoudre les problèmes avec votre manifeste</span><span class="sxs-lookup"><span data-stu-id="6168a-141">Use runtime logging to troubleshoot issues with your manifest</span></span>

<span data-ttu-id="6168a-142">Pour utiliser la journalisation runtime pour résoudre les problèmes de chargement d’un complément, procédez comme suit :</span><span class="sxs-lookup"><span data-stu-id="6168a-142">To use runtime logging to troubleshoot issues loading an add-in:</span></span>
 
1. <span data-ttu-id="6168a-143">[Chargez une version test de votre complément](sideload-office-add-ins-for-testing.md).</span><span class="sxs-lookup"><span data-stu-id="6168a-143">[Sideload your add-in](sideload-office-add-ins-for-testing.md) for testing.</span></span> 

    > [!NOTE]
    > <span data-ttu-id="6168a-144">Nous vous recommandons de charger uniquement une version test du complément que vous testez pour réduire le nombre de messages dans le fichier journal.</span><span class="sxs-lookup"><span data-stu-id="6168a-144">We recommend that you sideload only the add-in that you are testing to minimize the number of messages in the log file.</span></span>

2. <span data-ttu-id="6168a-145">Si rien ne se produit et que votre complément n’apparaît pas (et ne s’affiche pas dans la boîte de dialogue des compléments), ouvrez le fichier journal.</span><span class="sxs-lookup"><span data-stu-id="6168a-145">If nothing happens and you don't see your add-in (and it's not appearing in the add-ins dialog box), open the log file.</span></span>

3. <span data-ttu-id="6168a-p107">Recherchez le fichier journal pour l’ID de votre complément, que vous définissez dans votre manifeste. Dans le fichier journal, cet ID est intitulé `SolutionId`.</span><span class="sxs-lookup"><span data-stu-id="6168a-p107">Search the log file for your add-in ID, which you define in your manifest. In the log file, this ID is labeled `SolutionId`.</span></span> 

<span data-ttu-id="6168a-p108">Dans l’exemple suivant, le fichier journal identifie un contrôle qui pointe vers un fichier de ressources qui n’existe pas. Pour cet exemple, la correction consistera à corriger la faute de frappe dans le manifeste ou à ajouter la ressource manquante.</span><span class="sxs-lookup"><span data-stu-id="6168a-p108">In the following example, the log file identifies a control that points to a resource file that doesn't exist. For this example, the fix would be to correct the typo in the manifest or to add the missing resource.</span></span>

![Capture d’écran d’un fichier journal avec une entrée qui spécifie un ID de ressource qui est introuvable](http://i.imgur.com/f8bouLA.png) 

## <a name="known-issues-with-runtime-logging"></a><span data-ttu-id="6168a-151">Problèmes connus avec la journalisation runtime</span><span class="sxs-lookup"><span data-stu-id="6168a-151">Known issues with runtime logging</span></span>

<span data-ttu-id="6168a-p109">Vous pouvez afficher des messages dans le fichier journal qui sont source de confusion ou classés de façon incorrecte. Par exemple :</span><span class="sxs-lookup"><span data-stu-id="6168a-p109">You might see messages in the log file that are confusing or that are classified incorrectly. For example:</span></span>

- <span data-ttu-id="6168a-154">Le message `Medium Current host not in add-in's host list` suivi de `Unexpected Parsed manifest targeting different host` est classé incorrectement en tant qu’erreur.</span><span class="sxs-lookup"><span data-stu-id="6168a-154">The message `Medium Current host not in add-in's host list` followed by `Unexpected Parsed manifest targeting different host` is incorrectly classified as an error.</span></span>

- <span data-ttu-id="6168a-155">Si vous voyez le message `Unexpected Add-in is missing required manifest fields    DisplayName` et qu’il ne contient pas de SolutionId, l’erreur n’est probablement pas liée au complément que vous déboguez.</span><span class="sxs-lookup"><span data-stu-id="6168a-155">If you see the message `Unexpected Add-in is missing required manifest fields    DisplayName` and it doesn't contain a SolutionId, the error is most likely not related to the add-in you are debugging.</span></span> 

- <span data-ttu-id="6168a-p110">Tous les messages `Monitorable` sont des erreurs attendues du point de vue du système. Parfois, ils indiquent un problème avec votre manifeste, comme un élément mal orthographié qui a été ignoré, mais n’a pas provoqué l’échec du manifeste.</span><span class="sxs-lookup"><span data-stu-id="6168a-p110">Any `Monitorable` messages are expected errors from a system point of view. Sometimes they indicate an issue with your manifest, such as a misspelled element that was skipped but didn't cause the manifest to fail.</span></span> 

## <a name="see-also"></a><span data-ttu-id="6168a-158">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="6168a-158">See also</span></span>

- [<span data-ttu-id="6168a-159">Manifeste XML des compléments Office</span><span class="sxs-lookup"><span data-stu-id="6168a-159">Office Add-ins XML manifest</span></span>](../develop/add-in-manifests.md)
- [<span data-ttu-id="6168a-160">Valider le manifeste d’un complément Office</span><span class="sxs-lookup"><span data-stu-id="6168a-160">Validate an Office Add-in's manifest</span></span>](troubleshoot-manifest.md)
- [<span data-ttu-id="6168a-161">Vider le cache Office</span><span class="sxs-lookup"><span data-stu-id="6168a-161">Clear the Office cache</span></span>](clear-cache.md)
- [<span data-ttu-id="6168a-162">Chargement de la version test des compléments Office</span><span class="sxs-lookup"><span data-stu-id="6168a-162">Sideload Office Add-ins for testing</span></span>](sideload-office-add-ins-for-testing.md)
- [<span data-ttu-id="6168a-163">Débogage des compléments Office</span><span class="sxs-lookup"><span data-stu-id="6168a-163">Debug Office Add-ins</span></span>](debug-add-ins-using-f12-developer-tools-on-windows-10.md)