---
title: Valider un manifeste de complément Office
description: Découvrez la validation d’un manifeste de complément Office à l’aide du schéma XML ainsi que d’autres outils.
ms.date: 12/31/2019
localization_priority: Normal
ms.openlocfilehash: 9cd1c353d6f73decb5e39df96cf66da5912b8f9c
ms.sourcegitcommit: 6c7c98f085dd20f827e0c388e672993412944851
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/06/2020
ms.locfileid: "42554621"
---
# <a name="validate-an-office-add-ins-manifest"></a><span data-ttu-id="dce77-103">Valider un manifeste de complément Office</span><span class="sxs-lookup"><span data-stu-id="dce77-103">Validate an Office Add-in's manifest</span></span>

<span data-ttu-id="dce77-104">Vous souhaitez peut-être valider le fichier manifeste de votre complément pour vous assurer que celui-ci est correct et complet.</span><span class="sxs-lookup"><span data-stu-id="dce77-104">You may want to validate your add-in's manifest file to ensure that it's correct and complete.</span></span> <span data-ttu-id="dce77-105">La validation peut également identifier les problèmes à l’origine de l’erreur « Votre manifeste de complément n’est pas valide » lorsque vous essayez de charger une version test de votre complément.</span><span class="sxs-lookup"><span data-stu-id="dce77-105">Validation can also identify issues that are causing the error "Your add-in manifest is not valid" when you attempt to sideload your add-in.</span></span> <span data-ttu-id="dce77-106">Cet article décrit plusieurs méthodes de validation du fichier manifeste.</span><span class="sxs-lookup"><span data-stu-id="dce77-106">This article describes multiple ways to validate the manifest file.</span></span>

> [!NOTE]
> <span data-ttu-id="dce77-107">Pour en savoir plus sur l’utilisation de la journalisation de l’exécution pour résoudre des problèmes relatifs au manifeste de votre complément, consultez [Déboguer votre complément avec la journalisation de l’exécution](runtime-logging.md).</span><span class="sxs-lookup"><span data-stu-id="dce77-107">For details about using runtime logging to troubleshoot issues with your add-in's manifest, see [Debug your add-in with runtime logging](runtime-logging.md).</span></span>

## <a name="validate-your-manifest-with-the-yeoman-generator-for-office-add-ins"></a><span data-ttu-id="dce77-108">Valider votre manifeste avec le générateur Yeoman pour les compléments Office</span><span class="sxs-lookup"><span data-stu-id="dce77-108">Validate your manifest with the Yeoman generator for Office Add-ins</span></span>

<span data-ttu-id="dce77-109">Si vous avez utilisé [le générateur Yeoman pour les compléments Office](https://www.npmjs.com/package/generator-office) pour créer votre complément, vous pouvez également l’utiliser pour valider le fichier manifeste de votre projet.</span><span class="sxs-lookup"><span data-stu-id="dce77-109">If you used the [Yeoman generator for Office Add-ins](https://www.npmjs.com/package/generator-office) to create your add-in, you can also use it to validate your project's manifest file.</span></span> <span data-ttu-id="dce77-110">Exécutez la commande suivante dans le répertoire racine de votre projet :</span><span class="sxs-lookup"><span data-stu-id="dce77-110">Run the following command in the root directory of your project:</span></span>

```command&nbsp;line
npm run validate
```

![Gif animé qui montre le validateur Yo Office exécuté sur la ligne de commande et les résultats générés indiquant « Validation Passed » (validation réussie)](../images/yo-office-validator.gif)

> [!NOTE]
> <span data-ttu-id="dce77-112">Pour accéder à cette fonctionnalité, votre projet de complément doit être créé à l’aide du [générateur Yeoman pour les compléments Office](https://www.npmjs.com/package/generator-office) (version 1.1.17 ou ultérieure).</span><span class="sxs-lookup"><span data-stu-id="dce77-112">To have access to this functionality, your add-in project must have been created by using [Yeoman generator for Office Add-ins](https://www.npmjs.com/package/generator-office) version 1.1.17 or later.</span></span>

## <a name="validate-your-manifest-with-office-addin-manifest"></a><span data-ttu-id="dce77-113">Valider votre manifeste avec office-addin-manifest</span><span class="sxs-lookup"><span data-stu-id="dce77-113">Validate your manifest with office-addin-manifest</span></span>

<span data-ttu-id="dce77-114">Si vous n’avez pas utilisé [le générateur Yeoman pour les compléments Office](https://www.npmjs.com/package/generator-office) pour créer votre complément, vous pouvez valider le fichier manifeste à l’aide de [office-addin-manifest](https://www.npmjs.com/package/office-addin-manifest).</span><span class="sxs-lookup"><span data-stu-id="dce77-114">If you didn't use the [Yeoman generator for Office Add-ins](https://www.npmjs.com/package/generator-office) to create your add-in, you can validate the manifest by using [office-addin-manifest](https://www.npmjs.com/package/office-addin-manifest).</span></span>

1. <span data-ttu-id="dce77-115">Installez [Node.js](https://nodejs.org/download/).</span><span class="sxs-lookup"><span data-stu-id="dce77-115">Install [Node.js](https://nodejs.org/download/).</span></span>

2. <span data-ttu-id="dce77-116">Exécutez la commande suivante dans le répertoire racine de votre projet.</span><span class="sxs-lookup"><span data-stu-id="dce77-116">Run the following command in the root directory of your project.</span></span> <span data-ttu-id="dce77-117">Remplacez `MANIFEST_FILE` par le nom du fichier manifeste.</span><span class="sxs-lookup"><span data-stu-id="dce77-117">Replace `MANIFEST_FILE` with the name of the manifest file.</span></span>

    ```command&nbsp;line
    npx office-addin-manifest validate MANIFEST_FILE
    ```

    > [!NOTE]
    > <span data-ttu-id="dce77-118">Si elle s’exécute, la commande renvoie le message d’erreur « La syntaxe de la commande n’est pas valide »</span><span class="sxs-lookup"><span data-stu-id="dce77-118">If running this command results in the error message "The command syntax is not valid."</span></span> <span data-ttu-id="dce77-119">(étant donné que la commande `validate` n’est pas reconnue), exécutez la commande suivante pour valider le manifeste (en remplaçant `MANIFEST_FILE` par le nom du fichier manifeste) :</span><span class="sxs-lookup"><span data-stu-id="dce77-119">(because the `validate` command is not recognized), run the following command to validate the manifest (replacing `MANIFEST_FILE` with the name of the manifest file):</span></span> 
    >
    > `npx --ignore-existing office-addin-manifest validate MANIFEST_FILE`

## <a name="validate-your-manifest-against-the-xml-schema"></a><span data-ttu-id="dce77-120">Validez votre manifeste par rapport au schéma XML</span><span class="sxs-lookup"><span data-stu-id="dce77-120">Validate your manifest against the XML schema</span></span>

<span data-ttu-id="dce77-121">Vous pouvez valider le fichier manifeste par rapport aux fichiers de [définition de schéma XML (XSD)](/openspecs/office_file_formats/ms-owemxml/c6a06390-34b8-4b42-82eb-b28be12494a8).</span><span class="sxs-lookup"><span data-stu-id="dce77-121">You can validate the manifest file against the [XML Schema Definition (XSD)](/openspecs/office_file_formats/ms-owemxml/c6a06390-34b8-4b42-82eb-b28be12494a8) files.</span></span> <span data-ttu-id="dce77-122">Cela permet de s’assurer que le fichier manifeste suit le schéma approprié, y compris les espaces de noms pour les éléments que vous utilisez.</span><span class="sxs-lookup"><span data-stu-id="dce77-122">This will ensure that the manifest file follows the correct schema, including any namespaces for the elements you are using.</span></span> <span data-ttu-id="dce77-123">Si vous avez copié des éléments à partir d’autres exemples de manifestes, vérifiez par deux fois que vous avez également **inclus les espaces de noms appropriés**.</span><span class="sxs-lookup"><span data-stu-id="dce77-123">If you copied elements from other sample manifests double check that you also **include the appropriate namespaces**.</span></span> <span data-ttu-id="dce77-124">Pour ce faire, vous pouvez utiliser un outil de validation de schéma XML.</span><span class="sxs-lookup"><span data-stu-id="dce77-124">You can use an XML schema validation tool to perform this validation.</span></span>

### <a name="to-use-a-command-line-xml-schema-validation-tool-to-validate-your-manifest"></a><span data-ttu-id="dce77-125">Pour utiliser un outil de validation de schéma XML à ligne de commande pour valider votre manifeste</span><span class="sxs-lookup"><span data-stu-id="dce77-125">To use a command-line XML schema validation tool to validate your manifest</span></span>

1. <span data-ttu-id="dce77-126">Installez [tar](https://www.gnu.org/software/tar/) et [libxml](http://xmlsoft.org/FAQ.html), si vous ne l’avez pas déjà fait.</span><span class="sxs-lookup"><span data-stu-id="dce77-126">Install [tar](https://www.gnu.org/software/tar/) and [libxml](http://xmlsoft.org/FAQ.html), if you haven't already.</span></span>

2. <span data-ttu-id="dce77-p106">Exécutez la commande suivante. Remplacez `XSD_FILE` par le chemin d’accès au fichier XSD manifeste et `XML_FILE` par le chemin d’accès au fichier XML manifeste.</span><span class="sxs-lookup"><span data-stu-id="dce77-p106">Run the following command. Replace `XSD_FILE` with the path to the manifest XSD file, and replace `XML_FILE` with the path to the manifest XML file.</span></span>
    
    ```command&nbsp;line
    xmllint --noout --schema XSD_FILE XML_FILE
    ```

## <a name="see-also"></a><span data-ttu-id="dce77-129">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="dce77-129">See also</span></span>

- [<span data-ttu-id="dce77-130">Manifeste XML des compléments Office</span><span class="sxs-lookup"><span data-stu-id="dce77-130">Office Add-ins XML manifest</span></span>](../develop/add-in-manifests.md)
- [<span data-ttu-id="dce77-131">Vider le cache Office</span><span class="sxs-lookup"><span data-stu-id="dce77-131">Clear the Office cache</span></span>](clear-cache.md)
- [<span data-ttu-id="dce77-132">Déboguer votre complément avec la journalisation de l’exécution</span><span class="sxs-lookup"><span data-stu-id="dce77-132">Debug your add-in with runtime logging</span></span>](runtime-logging.md)
- [<span data-ttu-id="dce77-133">Chargement de la version test des compléments Office</span><span class="sxs-lookup"><span data-stu-id="dce77-133">Sideload Office Add-ins for testing</span></span>](sideload-office-add-ins-for-testing.md)
- [<span data-ttu-id="dce77-134">Débogage des compléments Office</span><span class="sxs-lookup"><span data-stu-id="dce77-134">Debug Office Add-ins</span></span>](debug-add-ins-using-f12-developer-tools-on-windows-10.md)