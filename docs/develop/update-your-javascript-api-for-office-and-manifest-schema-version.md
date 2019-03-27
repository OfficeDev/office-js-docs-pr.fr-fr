---
title: Télécharger la dernière version de l’API JavaScript pour la bibliothèque Office et la version 1.1 du schéma de manifeste du complément
description: Mettez à jour vos fichiers JavaScript (Office.js et fichiers .js propres aux applications) et le fichier de validation du manifeste du complément dans votre projet Complément Office vers la version 1.1.
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 7cbda821897b33a19e4bc9eeac27a096e01bc217
ms.sourcegitcommit: a2950492a2337de3180b713f5693fe82dbdd6a17
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/27/2019
ms.locfileid: "30872192"
---
# <a name="update-to-the-latest-javascript-api-for-office-library-and-version-11-add-in-manifest-schema"></a><span data-ttu-id="a6ae8-103">Télécharger la dernière version de l’API JavaScript pour la bibliothèque Office et la version 1.1 du schéma de manifeste du complément</span><span class="sxs-lookup"><span data-stu-id="a6ae8-103">Update to the latest JavaScript API for Office library and version 1.1 add-in manifest schema</span></span>

<span data-ttu-id="a6ae8-104">Cet article décrit comment mettre à jour vers la version 1.1 les fichiers JavaScript pour Office (Office.js et fichiers .js propres aux applications) et le fichier de validation du manifeste du complément utilisés dans votre projet de complément Office.</span><span class="sxs-lookup"><span data-stu-id="a6ae8-104">This article describes how to update your JavaScript files (Office.js and app-specific .js files) and add-in manifest validation file in your Office Add-in project to version 1.1.</span></span>

> [!NOTE]
> <span data-ttu-id="a6ae8-105">Les projets créés dans Visual Studio 2017 utiliseront déjà la version 1.1.</span><span class="sxs-lookup"><span data-stu-id="a6ae8-105">Projects created in Visual Studio 2017 will already use version 1.1.</span></span> <span data-ttu-id="a6ae8-106">Il existe toutefois des mises à jour mineures occasionnelles vers la version 1.1 que vous pouvez appliquer à l’aide des techniques décrites dans cet article.</span><span class="sxs-lookup"><span data-stu-id="a6ae8-106">However there are occasional minor updates to version 1.1 that you can apply by using the techniques in this article.</span></span>

## <a name="use-the-most-up-to-date-project-files"></a><span data-ttu-id="a6ae8-107">Utilisation des fichiers de projet les plus récents</span><span class="sxs-lookup"><span data-stu-id="a6ae8-107">Use the most up-to-date project files</span></span>

<span data-ttu-id="a6ae8-108">Si vous utilisez Visual Studio pour développer votre complément, et que vous souhaitez utiliser les [nouveaux membres d’API](/office/dev/add-ins/reference/what's-changed-in-the-javascript-api-for-office) de l’interface API JavaScript pour Office et les [fonctionnalités de la version 1.1 du manifeste du complément](../develop/add-in-manifests.md) (qui est validé par rapport à offappmanifest-1.1.xsd), vous devez télécharger Visual Studio 2017.</span><span class="sxs-lookup"><span data-stu-id="a6ae8-108">If you use Visual Studio to develop your add-in, to use the [newest API members](/office/dev/add-ins/reference/what's-changed-in-the-javascript-api-for-office) of the JavaScript API for Office and the [v1.1 features of the add-in manifest](../develop/add-in-manifests.md) (which is validated against offappmanifest-1.1.xsd), you need to download Visual Studio 2017.</span></span> <span data-ttu-id="a6ae8-109">Pour télécharger Visual Studio 2017, consultez la [Page Visual Studio IDE](https://visualstudio.microsoft.com/vs/).</span><span class="sxs-lookup"><span data-stu-id="a6ae8-109">To download Visual Studio 2017, see the [Visual Studio IDE page](https://visualstudio.microsoft.com/vs/).</span></span> <span data-ttu-id="a6ae8-110">Lors de l’installation, vous devez sélectionner la charge de travail de développement Office/SharePoint.</span><span class="sxs-lookup"><span data-stu-id="a6ae8-110">During installation you'll need to select the Office/SharePoint development workload.</span></span>

<span data-ttu-id="a6ae8-111">Si vous utilisez un éditeur de texte ou une interface IDE autre que Visual Studio pour développer votre complément, vous devez mettre à jour les références vers le CDN pour Office.js et la version de schéma référencée dans le manifeste de votre complément.</span><span class="sxs-lookup"><span data-stu-id="a6ae8-111">If you use a text editor or IDE other than Visual Studio to develop your add-in, you need to update the references to the CDN for Office.js and the version of schema referenced in your add-in's manifest.</span></span>

<span data-ttu-id="a6ae8-112">Pour exécuter un complément développé à l’aide des fonctionnalités nouvelles et mises à jour du manifeste du complément et de l’API d’Office.js, vos clients doivent exécuter des produits locaux Office 2013 SP1 ou version ultérieure, et le cas échéant, SharePoint Server 2013 SP1 et des produits serveur associés, Exchange Server 2013 Service Pack 1 (SP1) ou des produits hébergés en ligne équivalents : Office 365, SharePoint Online et Exchange Online.</span><span class="sxs-lookup"><span data-stu-id="a6ae8-112">To run an add-in developed using new and updated Office.js API and add-in manifest features, your customers must be running Office 2013 SP1 or later version on-premises products, and where applicable, SharePoint Server 2013 SP1 and related server products, Exchange Server 2013 Service Pack 1 (SP1), or the equivalent online hosted products: Office 365, SharePoint Online, and Exchange Online.</span></span>

<span data-ttu-id="a6ae8-113">Pour télécharger des produits Office, SharePoint et Exchange SP1, voir :</span><span class="sxs-lookup"><span data-stu-id="a6ae8-113">To download Office, SharePoint, and Exchange SP1 products, see the following:</span></span>

- [<span data-ttu-id="a6ae8-114">Liste de toutes les mises à jour Service Pack 1 (SP1) pour Microsoft Office 2013 et les produits bureautiques connexes</span><span class="sxs-lookup"><span data-stu-id="a6ae8-114">List of all Service Pack 1 (SP1) updates for Microsoft Office 2013 and related desktop products</span></span>](https://support.microsoft.com/kb/2850036)

- [<span data-ttu-id="a6ae8-115">Liste de toutes les mises à jour Service Pack 1 (SP1) pour Microsoft SharePoint Server 2013 et les produits serveur connexes</span><span class="sxs-lookup"><span data-stu-id="a6ae8-115">List of all Service Pack 1 (SP1) updates for Microsoft SharePoint Server 2013 and related server products</span></span>](https://support.microsoft.com/kb/2850035)

- [<span data-ttu-id="a6ae8-116">Description du Service Pack 1 d’Exchange Server 2013</span><span class="sxs-lookup"><span data-stu-id="a6ae8-116">Description of Exchange Server 2013 Service Pack 1</span></span>](https://support.microsoft.com/kb/2926248)


## <a name="updating-an-office-add-in-project-created-with-visual-studio"></a><span data-ttu-id="a6ae8-117">Mise à jour d’un projet de complément Office créé avec Visual Studio</span><span class="sxs-lookup"><span data-stu-id="a6ae8-117">Updating an Office Add-in project created with Visual Studio</span></span>

<span data-ttu-id="a6ae8-118">Pour les projets créés avant la sortie de la version 1.1 de l’API JavaScript pour Office et du schéma de manifeste de complément, vous pouvez mettre à jour les fichiers d’un projet à l’aide du  **gestionnaire de package NuGet**, puis mettre à jour les pages HTML de votre complément pour les référencer.</span><span class="sxs-lookup"><span data-stu-id="a6ae8-118">For projects created before the release of v1.1 of the JavaScript API for Office and add-in manifest schema, you can update a project's files using the  **NuGet Package Manager**, and then update your add-in's HTML pages to reference them.</span></span> 

<span data-ttu-id="a6ae8-119">Notez que le processus de mise à jour est appliqué  _par projet_  ; vous devrez répéter le processus de mise à jour pour chaque projet de complément dans lequel vous souhaitez utiliser la version 1.1 d’Office.js et du schéma de manifeste de complément.</span><span class="sxs-lookup"><span data-stu-id="a6ae8-119">Note that the update process is applied on a  _per-project basis_ - you'll need to repeat the updating process for each add-in project in which you want to use v1.1 of Office.js and add-in manifest schema.</span></span>


### <a name="update-the-javascript-api-for-office-library-files-in-your-project-to-the-newest-release"></a><span data-ttu-id="a6ae8-120">Mise à jour des fichiers de bibliothèque de l’API JavaScript pour Office dans votre projet vers la dernière version</span><span class="sxs-lookup"><span data-stu-id="a6ae8-120">Update the JavaScript API for Office library files in your project to the newest release</span></span>
<span data-ttu-id="a6ae8-121">Les étapes suivantes permettent de mettre à jour vos fichiers de bibliothèque Office vers la dernière version.</span><span class="sxs-lookup"><span data-stu-id="a6ae8-121">The following steps will update your Office library files to the latest version.</span></span> <span data-ttu-id="a6ae8-122">Les étapes utilisent Visual Studio 2017, mais elles sont similaires pour Visual Studio 2015.</span><span class="sxs-lookup"><span data-stu-id="a6ae8-122">The steps use Visual Studio 2017, but they are similar for Visual Studio 2015.</span></span>

1. <span data-ttu-id="a6ae8-123">Dans Visual Studio 2017, ouvrez ou créez un projet de **complément Office**.</span><span class="sxs-lookup"><span data-stu-id="a6ae8-123">In Visual Studio 2017, open or create a new  **Office Add-in** project.</span></span>
2. <span data-ttu-id="a6ae8-124">Sélectionnez **Outils**  >  **Gestionnaire de package NuGet**  >  **Gérer les packages NuGet pour la solution**.</span><span class="sxs-lookup"><span data-stu-id="a6ae8-124">Choose  **Tools** > **NuGet Package Manager** > **Manage Nuget Packages for Solution**.</span></span>
3. <span data-ttu-id="a6ae8-125">Dans **Gestionnaire de package NuGet**, sélectionnez **nuget.org** pour **Source du package**.</span><span class="sxs-lookup"><span data-stu-id="a6ae8-125">In the  **NuGet Package Manager**, select  **nuget.org** for **Package source**.</span></span>
4. <span data-ttu-id="a6ae8-126">Sélectionnez l’onglet **Mises à jour**.</span><span class="sxs-lookup"><span data-stu-id="a6ae8-126">Choose the **Updates** tab.</span></span>
5. <span data-ttu-id="a6ae8-127">Sélectionnez Microsoft.Office.js.</span><span class="sxs-lookup"><span data-stu-id="a6ae8-127">Select Microsoft.Office.js.</span></span>
6. <span data-ttu-id="a6ae8-128">Dans le volet de gauche, sélectionnez **Mettre à jour** et terminez le processus de mise à jour du package.</span><span class="sxs-lookup"><span data-stu-id="a6ae8-128">In the left pane, choose **Update** and complete the package update process.</span></span>

<span data-ttu-id="a6ae8-129">Vous devez effectuer quelques étapes supplémentaires pour terminer la mise à jour.</span><span class="sxs-lookup"><span data-stu-id="a6ae8-129">You'll need to take a few additional steps to complete the update.</span></span> <span data-ttu-id="a6ae8-130">Dans la balise **head** des pages HTML de votre complément, commentez ou supprimez toute référence de script office.js existante, puis référencez la bibliothèque mise à jour de l’API JavaScript pour Office de la manière suivante :</span><span class="sxs-lookup"><span data-stu-id="a6ae8-130">In the **head** tag of your add-in's HTML pages, comment out or delete any existing office.js script references, and reference the updated JavaScript API for Office library as follows:</span></span>

  ```html
  <script src="https://appsforoffice.microsoft.com/lib/1/hosted/Office.js" type="text/javascript"></script>
  ```

   > [!NOTE] 
   > <span data-ttu-id="a6ae8-131">La valeur `/1/` de `office.js` dans l’URL du CDN préconise l’utilisation de la dernière version incrémentielle comprise dans la version 1 d’Office.js.</span><span class="sxs-lookup"><span data-stu-id="a6ae8-131">The `/1/` in the `office.js` in the CDN URL specifies to use the latest incremental release within version 1 of Office.js.</span></span>


### <a name="update-the-manifest-file-in-your-project-to-use-schema-version-11"></a><span data-ttu-id="a6ae8-132">Mise à jour du fichier manifeste dans votre projet afin d’utiliser la version 1.1 du schéma</span><span class="sxs-lookup"><span data-stu-id="a6ae8-132">Update the manifest file in your project to use schema version 1.1</span></span>

<span data-ttu-id="a6ae8-133">Dans le fichier manifeste de votre complément, mettez à jour l’attribut **xmlns** de l’élément **OfficeApp** en appliquant la valeur `1.1` à la version (sans modifier les attributs autres que **xmlns**).</span><span class="sxs-lookup"><span data-stu-id="a6ae8-133">In your add-in's manifest file, update the **xmlns** attribute of the **OfficeApp** element changing the version value to `1.1` (leaving attributes other than the **xmlns** attribute unchanged).</span></span>

```xml
<?xml version="1.0" encoding="utf-8"?>
<OfficeApp xsi:type="ContentApp"
    xmlns:xsi="https://www.w3.org/2001/XMLSchema-instance"
    xmlns="http://schemas.microsoft.com/office/appforoffice/1.1">
  
  <!-- manifest contents -->

</OfficeApp>
```

> [!NOTE]
> <span data-ttu-id="a6ae8-134">Une fois le schéma manifeste du complément mis à jour vers la version 1.1, vous devez supprimer les éléments **Capabilities** et **Capability**, et les remplacer par les éléments [Hosts](/office/dev/add-ins/reference/manifest/hosts) et [Host](/office/dev/add-ins/reference/manifest/host) ou [Requirements et Requirement](specify-office-hosts-and-api-requirements.md).</span><span class="sxs-lookup"><span data-stu-id="a6ae8-134">After updating the version of the add-in manifest schema to 1.1, you will need to remove the  **Capabilities** and **Capability** elements, and replace them with either the [Hosts](/office/dev/add-ins/reference/manifest/hosts) and [Host](/office/dev/add-ins/reference/manifest/host) elements or the [Requirements and Requirement elements](specify-office-hosts-and-api-requirements.md).</span></span>

## <a name="updating-an-office-add-in-project-created-with-a-text-editor-or-other-ide"></a><span data-ttu-id="a6ae8-135">Mise à jour d’un projet de complément Office créé avec un éditeur de texte ou une autre IDE</span><span class="sxs-lookup"><span data-stu-id="a6ae8-135">Updating an Office Add-in project created with a text editor or other IDE</span></span>

<span data-ttu-id="a6ae8-136">Pour les projets créés avant la publication de la version 1.1 de l’API JavaScript pour Office et du schéma de manifeste de complément, vous devez mettre à jour les pages HTML de votre complément afin de faire référence au CDN de la bibliothèque version 1.1, ainsi que mettre à jour le fichier de manifeste de votre complément pour utiliser le schéma version 1.1.</span><span class="sxs-lookup"><span data-stu-id="a6ae8-136">For projects created before the release of v1.1 of the JavaScript API for Office and add-in manifest schema, you need to update your add-in's HTML pages to reference CDN of the v1.1 library, and update your add-in's manifest file to use schema v1.1.</span></span> 

<span data-ttu-id="a6ae8-137">Le processus de mise à jour est appliqué  _par projet_  ; vous devrez répéter le processus de mise à jour pour chaque projet de complément dans lequel vous souhaitez utiliser la version 1.1 d’Office.js et du schéma de manifeste de complément.</span><span class="sxs-lookup"><span data-stu-id="a6ae8-137">The update process is applied on a  _per-project basis_ - you'll need to repeat the updating process for each add-in project in which you want to use v1.1 of Office.js and add-in manifest schema.</span></span>

<span data-ttu-id="a6ae8-138">Vous n’avez pas besoin de copies locales des fichiers de l’interface API JavaScript pour Office (fichiers Office.js et fichiers .js propres aux applications) pour développer un complément Office (si vous référencez le CDN pour Office.js, les fichiers requis sont téléchargés lors de l’exécution), mais si vous voulez une copie locale des fichiers de bibliothèque, vous pouvez utiliser l’[utilitaire de ligne de commande NuGet](https://docs.nuget.org/consume/installing-nuget) et la commande `Install-Package Microsoft.Office.js` pour les télécharger.</span><span class="sxs-lookup"><span data-stu-id="a6ae8-138">You don't need local copies of the JavaScript API for Office files (Office.js and app-specific .js files) to develop anOffice Add-in (referencing the CDN for Office.js downloads the necessary files at runtime), but if you want a local copy of the library files you can use the [NuGet Command-Line Utility](https://docs.nuget.org/consume/installing-nuget) and the `Install-Package Microsoft.Office.js` command to download them.</span></span>

> [!NOTE]
> <span data-ttu-id="a6ae8-139">pour obtenir une copie du fichier XSD (définition du schéma XML) pour le manifeste de complément version 1.1, consultez les [informations de référence sur le schéma des manifestes des compléments Office (version 1.1)](../develop/add-in-manifests.md).</span><span class="sxs-lookup"><span data-stu-id="a6ae8-139">To get a copy of the XSD (XML Schema Definition) for the v1.1 add-in manifest, see the listing in [Schema reference for Office Add-ins manifests (v1.1)](../develop/add-in-manifests.md).</span></span>


### <a name="update-the-javascript-api-for-office-library-files-in-your-project-to-use-the-newest-release"></a><span data-ttu-id="a6ae8-140">Mise à jour des fichiers de bibliothèque de l’API JavaScript pour Office dans votre projet pour utiliser la dernière version</span><span class="sxs-lookup"><span data-stu-id="a6ae8-140">Update the JavaScript API for Office library files in your project to use the newest release</span></span>

1. <span data-ttu-id="a6ae8-141">Ouvrez les pages HTML de votre complément dans un éditeur de texte ou une interface IDE.</span><span class="sxs-lookup"><span data-stu-id="a6ae8-141">Open the HTML pages for your add-in in your text editor or IDE.</span></span>

2. <span data-ttu-id="a6ae8-142">Dans la balise **head** des pages HTML de votre complément, commentez ou supprimez toute référence de script office.js existante, puis référencez la bibliothèque mise à jour de l’API JavaScript pour Office de la manière suivante :</span><span class="sxs-lookup"><span data-stu-id="a6ae8-142">In the **head** tag of your add-in's HTML pages, comment out or delete any existing office.js script references, and reference the updated JavaScript API for Office library as follows:</span></span>

    ```html
    <script src="https://appsforoffice.microsoft.com/lib/1/hosted/Office.js" type="text/javascript"></script>
    ```

   > [!NOTE]
   > <span data-ttu-id="a6ae8-143">la valeur `/1/` devant `office.js` dans l’URL du CDN préconise l’utilisation de la dernière version incrémentielle comprise dans la version 1 d’Office.js.</span><span class="sxs-lookup"><span data-stu-id="a6ae8-143">The `/1/` in front of `office.js` in the CDN URL specifies to use the latest incremental release within version 1 of Office.js.</span></span>

### <a name="update-the-manifest-file-in-your-project-to-use-schema-version-11"></a><span data-ttu-id="a6ae8-144">Mise à jour du fichier manifeste dans votre projet afin d’utiliser la version 1.1 du schéma</span><span class="sxs-lookup"><span data-stu-id="a6ae8-144">Update the manifest file in your project to use schema version 1.1</span></span>

<span data-ttu-id="a6ae8-145">Dans le fichier manifeste de votre complément, mettez à jour l’attribut **xmlns** de l’élément **OfficeApp** en appliquant la valeur `1.1` à la version (sans modifier les attributs autres que **xmlns**).</span><span class="sxs-lookup"><span data-stu-id="a6ae8-145">In your add-in's manifest file, update the **xmlns** attribute of the **OfficeApp** element changing the version value to `1.1` (leaving attributes other than the **xmlns** attribute unchanged).</span></span>

```xml
<?xml version="1.0" encoding="utf-8"?>
<OfficeApp xsi:type="ContentApp"
  xmlns:xsi="https://www.w3.org/2001/XMLSchema-instance"
  xmlns="http://schemas.microsoft.com/office/appforoffice/1.1">
  
  <!-- manifest contents -->

</OfficeApp>
```

> [!NOTE]
> <span data-ttu-id="a6ae8-146">Une fois le schéma manifeste du complément mis à jour vers la version 1.1, vous devez supprimer les éléments **Capabilities** et **Capability**, et les remplacer par les éléments [Hosts](/office/dev/add-ins/reference/manifest/hosts) et [Host](/office/dev/add-ins/reference/manifest/host) ou [Requirements et Requirement](specify-office-hosts-and-api-requirements.md).</span><span class="sxs-lookup"><span data-stu-id="a6ae8-146">After updating the version of the add-in manifest schema to 1.1, you will need to remove the  **Capabilities** and **Capability** elements, and replace them with either the [Hosts](/office/dev/add-ins/reference/manifest/hosts) and [Host](/office/dev/add-ins/reference/manifest/host) elements or the [Requirements and Requirement elements](specify-office-hosts-and-api-requirements.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="a6ae8-147">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="a6ae8-147">See also</span></span>

- <span data-ttu-id="a6ae8-148">[Spécification des exigences en matière d'hôtes Office et d'API](specify-office-hosts-and-api-requirements.md) ]</span><span class="sxs-lookup"><span data-stu-id="a6ae8-148">[Specify Office hosts and API requirements](specify-office-hosts-and-api-requirements.md) ]</span></span>
- [<span data-ttu-id="a6ae8-149">Présentation de l’API JavaScript pour Office</span><span class="sxs-lookup"><span data-stu-id="a6ae8-149">Understanding the JavaScript API for Office</span></span>](understanding-the-javascript-api-for-office.md)
- [<span data-ttu-id="a6ae8-150">Interface API JavaScript pour Office</span><span class="sxs-lookup"><span data-stu-id="a6ae8-150">JavaScript API for Office</span></span>](/office/dev/add-ins/reference/javascript-api-for-office)
- [<span data-ttu-id="a6ae8-151">Informations de référence sur le schéma des manifestes des applications pour Office (version 1.1)</span><span class="sxs-lookup"><span data-stu-id="a6ae8-151">Schema reference for Office Add-ins manifests (v1.1)</span></span>](../develop/add-in-manifests.md)
