---
title: Valider un manifeste de complément Office
description: Découvrez comment valider le manifeste d’un complément Office à l’aide du schéma XML et d’autres outils.
ms.date: 04/16/2020
localization_priority: Normal
ms.openlocfilehash: fee4fd048092734eb479f1993c69fcf99c153c79
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/08/2020
ms.locfileid: "44611098"
---
# <a name="validate-an-office-add-ins-manifest"></a><span data-ttu-id="486f7-103">Valider un manifeste de complément Office</span><span class="sxs-lookup"><span data-stu-id="486f7-103">Validate an Office Add-in's manifest</span></span>

<span data-ttu-id="486f7-104">Vous souhaitez peut-être valider le fichier manifeste de votre complément pour vous assurer que celui-ci est correct et complet.</span><span class="sxs-lookup"><span data-stu-id="486f7-104">You may want to validate your add-in's manifest file to ensure that it's correct and complete.</span></span> <span data-ttu-id="486f7-105">La validation peut également identifier les problèmes à l’origine de l’erreur « Votre manifeste de complément n’est pas valide » lorsque vous essayez de charger une version test de votre complément.</span><span class="sxs-lookup"><span data-stu-id="486f7-105">Validation can also identify issues that are causing the error "Your add-in manifest is not valid" when you attempt to sideload your add-in.</span></span> <span data-ttu-id="486f7-106">Cet article décrit plusieurs méthodes de validation du fichier manifeste.</span><span class="sxs-lookup"><span data-stu-id="486f7-106">This article describes multiple ways to validate the manifest file.</span></span>

> [!NOTE]
> <span data-ttu-id="486f7-107">Pour en savoir plus sur l’utilisation de la journalisation de l’exécution pour résoudre des problèmes relatifs au manifeste de votre complément, consultez [Déboguer votre complément avec la journalisation de l’exécution](runtime-logging.md).</span><span class="sxs-lookup"><span data-stu-id="486f7-107">For details about using runtime logging to troubleshoot issues with your add-in's manifest, see [Debug your add-in with runtime logging](runtime-logging.md).</span></span>

## <a name="validate-your-manifest-with-the-yeoman-generator-for-office-add-ins"></a><span data-ttu-id="486f7-108">Valider votre manifeste avec le générateur Yeoman pour les compléments Office</span><span class="sxs-lookup"><span data-stu-id="486f7-108">Validate your manifest with the Yeoman generator for Office Add-ins</span></span>

<span data-ttu-id="486f7-109">Si vous avez utilisé [le générateur Yeoman pour les compléments Office](https://www.npmjs.com/package/generator-office) pour créer votre complément, vous pouvez également l’utiliser pour valider le fichier manifeste de votre projet.</span><span class="sxs-lookup"><span data-stu-id="486f7-109">If you used the [Yeoman generator for Office Add-ins](https://www.npmjs.com/package/generator-office) to create your add-in, you can also use it to validate your project's manifest file.</span></span> <span data-ttu-id="486f7-110">Exécutez la commande suivante dans le répertoire racine de votre projet :</span><span class="sxs-lookup"><span data-stu-id="486f7-110">Run the following command in the root directory of your project:</span></span>

```command&nbsp;line
npm run validate
```

![Gif animé qui montre le validateur Yo Office exécuté sur la ligne de commande et les résultats générés indiquant « Validation Passed » (validation réussie)](../images/yo-office-validator.gif)

> [!NOTE]
> <span data-ttu-id="486f7-112">Pour accéder à cette fonctionnalité, votre projet de complément doit être créé à l’aide du [générateur Yeoman pour les compléments Office](https://www.npmjs.com/package/generator-office) (version 1.1.17 ou ultérieure).</span><span class="sxs-lookup"><span data-stu-id="486f7-112">To have access to this functionality, your add-in project must have been created by using [Yeoman generator for Office Add-ins](https://www.npmjs.com/package/generator-office) version 1.1.17 or later.</span></span>

## <a name="validate-your-manifest-with-office-addin-manifest"></a><span data-ttu-id="486f7-113">Valider votre manifeste avec office-addin-manifest</span><span class="sxs-lookup"><span data-stu-id="486f7-113">Validate your manifest with office-addin-manifest</span></span>

<span data-ttu-id="486f7-114">Si vous n’avez pas utilisé [le générateur Yeoman pour les compléments Office](https://www.npmjs.com/package/generator-office) pour créer votre complément, vous pouvez valider le fichier manifeste à l’aide de [office-addin-manifest](https://www.npmjs.com/package/office-addin-manifest).</span><span class="sxs-lookup"><span data-stu-id="486f7-114">If you didn't use the [Yeoman generator for Office Add-ins](https://www.npmjs.com/package/generator-office) to create your add-in, you can validate the manifest by using [office-addin-manifest](https://www.npmjs.com/package/office-addin-manifest).</span></span>

1. <span data-ttu-id="486f7-115">Installez [Node.js](https://nodejs.org/download/).</span><span class="sxs-lookup"><span data-stu-id="486f7-115">Install [Node.js](https://nodejs.org/download/).</span></span>

2. <span data-ttu-id="486f7-116">Exécutez la commande suivante dans le répertoire racine de votre projet.</span><span class="sxs-lookup"><span data-stu-id="486f7-116">Run the following command in the root directory of your project.</span></span> 

    ```command&nbsp;line
    npm run validate
    ```

    > [!NOTE]
    > <span data-ttu-id="486f7-117">Si cette commande n’est pas disponible ou ne fonctionne pas, exécutez la commande suivante pour forcer l’utilisation de la dernière version de l’outil Office-AddIn-manifest ( `MANIFEST_FILE` à remplacer par le nom du fichier manifeste) :</span><span class="sxs-lookup"><span data-stu-id="486f7-117">If this command is not available or not working, run the following command instead to force the use of the latest version of the office-addin-manifest tool (replacing `MANIFEST_FILE` with the name of the manifest file):</span></span>
    >
    > ```command&nbsp;line
    > npx --ignore-existing office-addin-manifest validate MANIFEST_FILE
    > ```

## <a name="validate-your-manifest-against-the-xml-schema"></a><span data-ttu-id="486f7-118">Validez votre manifeste par rapport au schéma XML</span><span class="sxs-lookup"><span data-stu-id="486f7-118">Validate your manifest against the XML schema</span></span>

<span data-ttu-id="486f7-119">Vous pouvez valider le fichier manifeste par rapport aux fichiers de [définition de schéma XML (XSD)](/openspecs/office_file_formats/ms-owemxml/c6a06390-34b8-4b42-82eb-b28be12494a8).</span><span class="sxs-lookup"><span data-stu-id="486f7-119">You can validate the manifest file against the [XML Schema Definition (XSD)](/openspecs/office_file_formats/ms-owemxml/c6a06390-34b8-4b42-82eb-b28be12494a8) files.</span></span> <span data-ttu-id="486f7-120">Cela permet de s’assurer que le fichier manifeste suit le schéma approprié, y compris les espaces de noms pour les éléments que vous utilisez.</span><span class="sxs-lookup"><span data-stu-id="486f7-120">This will ensure that the manifest file follows the correct schema, including any namespaces for the elements you are using.</span></span> <span data-ttu-id="486f7-121">Si vous avez copié des éléments à partir d’autres exemples de manifestes, vérifiez par deux fois que vous avez également **inclus les espaces de noms appropriés**.</span><span class="sxs-lookup"><span data-stu-id="486f7-121">If you copied elements from other sample manifests double check that you also **include the appropriate namespaces**.</span></span> <span data-ttu-id="486f7-122">Pour ce faire, vous pouvez utiliser un outil de validation de schéma XML.</span><span class="sxs-lookup"><span data-stu-id="486f7-122">You can use an XML schema validation tool to perform this validation.</span></span>

### <a name="to-use-a-command-line-xml-schema-validation-tool-to-validate-your-manifest"></a><span data-ttu-id="486f7-123">Pour utiliser un outil de validation de schéma XML à ligne de commande pour valider votre manifeste</span><span class="sxs-lookup"><span data-stu-id="486f7-123">To use a command-line XML schema validation tool to validate your manifest</span></span>

1. <span data-ttu-id="486f7-124">Installez [tar](https://www.gnu.org/software/tar/) et [libxml](http://xmlsoft.org/FAQ.html), si vous ne l’avez pas déjà fait.</span><span class="sxs-lookup"><span data-stu-id="486f7-124">Install [tar](https://www.gnu.org/software/tar/) and [libxml](http://xmlsoft.org/FAQ.html), if you haven't already.</span></span>

2. <span data-ttu-id="486f7-p104">Exécutez la commande suivante. Remplacez `XSD_FILE` par le chemin d’accès au fichier XSD manifeste et `XML_FILE` par le chemin d’accès au fichier XML manifeste.</span><span class="sxs-lookup"><span data-stu-id="486f7-p104">Run the following command. Replace `XSD_FILE` with the path to the manifest XSD file, and replace `XML_FILE` with the path to the manifest XML file.</span></span>
    
    ```command&nbsp;line
    xmllint --noout --schema XSD_FILE XML_FILE
    ```

## <a name="see-also"></a><span data-ttu-id="486f7-127">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="486f7-127">See also</span></span>

- [<span data-ttu-id="486f7-128">Manifeste XML des compléments Office</span><span class="sxs-lookup"><span data-stu-id="486f7-128">Office Add-ins XML manifest</span></span>](../develop/add-in-manifests.md)
- [<span data-ttu-id="486f7-129">Vider le cache Office</span><span class="sxs-lookup"><span data-stu-id="486f7-129">Clear the Office cache</span></span>](clear-cache.md)
- [<span data-ttu-id="486f7-130">Déboguer votre complément avec la journalisation de l’exécution</span><span class="sxs-lookup"><span data-stu-id="486f7-130">Debug your add-in with runtime logging</span></span>](runtime-logging.md)
- [<span data-ttu-id="486f7-131">Chargement de la version test des compléments Office</span><span class="sxs-lookup"><span data-stu-id="486f7-131">Sideload Office Add-ins for testing</span></span>](sideload-office-add-ins-for-testing.md)
- [<span data-ttu-id="486f7-132">Débogage des compléments Office</span><span class="sxs-lookup"><span data-stu-id="486f7-132">Debug Office Add-ins</span></span>](debug-add-ins-using-f12-developer-tools-on-windows-10.md)
