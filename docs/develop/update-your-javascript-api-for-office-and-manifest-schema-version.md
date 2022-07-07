---
title: Mise à jour vers la dernière bibliothèque d’API JavaScript Office et le schéma de manifeste du complément version 1.1
description: Mettez à jour vos fichiers JavaScript (Office.js et fichiers .js propres aux applications) et le fichier de validation du manifeste du complément dans votre projet Complément Office vers la version 1.1.
ms.date: 01/14/2022
ms.localizationpriority: medium
ms.openlocfilehash: 32fcadb6a36ca540a799f8d6a5dfa671ee5e5de8
ms.sourcegitcommit: 4ba5f750358c139c93eb2170ff2c97322dfb50df
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 07/06/2022
ms.locfileid: "66660199"
---
# <a name="update-to-the-latest-office-javascript-api-library-and-version-11-add-in-manifest-schema"></a>Mise à jour vers la dernière bibliothèque d’API JavaScript Office et le schéma de manifeste du complément version 1.1

Cet article décrit comment mettre à jour vers la version 1.1 les fichiers JavaScript pour Office (Office.js et fichiers .js propres aux applications) et le fichier de validation du manifeste du complément utilisés dans votre projet de complément Office.

> [!NOTE]
> Les projets créés dans Visual Studio 2019 utilisent déjà la version 1.1. Il existe toutefois des mises à jour mineures occasionnelles vers la version 1.1 que vous pouvez appliquer à l’aide des techniques décrites dans cet article.

## <a name="use-the-most-up-to-date-project-files"></a>Utilisation des fichiers de projet les plus récents

Si vous utilisez Visual Studio pour développer votre complément, pour utiliser les membres d’API les plus récents de l’API JavaScript Office et les [fonctionnalités v1.1 du manifeste de complément](../develop/add-in-manifests.md) (qui est validé par rapport à offappmanifest-1.1.xsd), vous devez télécharger Visual Studio 2019. Pour télécharger Visual Studio 2019, consultez la [page IDE de Visual Studio](https://visualstudio.microsoft.com/vs/). Lors de l’installation, vous devez sélectionner la charge de travail de développement Office/SharePoint.

Si vous utilisez un éditeur de texte ou un IDE autre que Visual Studio pour développer votre complément, vous devez mettre à jour les références au réseau de distribution de contenu (CDN) pour Office.js et la version du schéma référencée dans le manifeste de votre complément.

Pour exécuter un complément développé à l’aide de fonctionnalités de manifeste d’API et de complément Office.js nouvelles et mises à jour, vos clients doivent exécuter des produits locaux Office 2013 SP1 ou version ultérieure, et, le cas échéant, SharePoint Server 2013 SP1 et les produits serveur associés, Exchange Server 2013 Service Pack 1 (SP1) ou les produits hébergés en ligne équivalents : Microsoft 365, SharePoint Online et Exchange Online.

Pour télécharger des produits Office, SharePoint et Exchange SP1, voir :

- [Liste de toutes les mises à jour Service Pack 1 (SP1) pour Microsoft Office 2013 et les produits bureautiques connexes](https://support.microsoft.com/kb/2850036)

- [Liste de toutes les mises à jour Service Pack 1 (SP1) pour Microsoft SharePoint Server 2013 et les produits serveur connexes](https://support.microsoft.com/kb/2850035)

- [Description du Service Pack 1 d’Exchange Server 2013](https://support.microsoft.com/kb/2926248)

## <a name="updating-an-office-add-in-project-created-with-visual-studio"></a>Mise à jour d’un projet de complément Office créé avec Visual Studio

Pour les projets créés avant la publication de la version 1.1 de l’API JavaScript Office et du schéma de manifeste de complément, vous pouvez mettre à jour les fichiers d’un projet à l’aide du **Gestionnaire de package NuGet**, puis mettre à jour les pages HTML de votre complément pour les référencer.

Notez que le processus de mise à jour est appliqué  _par projet_  ; vous devrez répéter le processus de mise à jour pour chaque projet de complément dans lequel vous souhaitez utiliser la version 1.1 d’Office.js et du schéma de manifeste de complément.

### <a name="update-the-office-javascript-api-library-files-in-your-project-to-the-newest-release"></a>Mettez à jour les fichiers de bibliothèque d’API JavaScript Office dans votre projet vers la version la plus récente

Les étapes suivantes mettent à jour vos fichiers de bibliothèque Office.js vers la dernière version. Les étapes utilisent Visual Studio 2019, mais elles sont similaires pour les versions précédentes de Visual Studio.

1. Dans Visual Studio 2019, ouvrez ou **créez un projet de complément Office** .
2. Choisissez **Outils** > **NuGet Package Manager** > **Gérer les packages Nuget pour la solution**.
3. Sélectionnez l’onglet **Mises à jour**.
4. Sélectionnez Microsoft.Office.js. Vérifiez que la source du package provient de **nuget.org**.
5. Dans le volet gauche, choisissez **Installer** et terminez le processus de mise à jour du package.

Vous devez effectuer quelques étapes supplémentaires pour terminer la mise à jour. Dans la balise **principale** des pages HTML de votre complément, commentez ou supprimez les références de script office.js existantes, puis référencez la bibliothèque d’API JavaScript Office mise à jour comme suit :

  ```html
  <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js" type="text/javascript"></script>
  ```

   > [!NOTE]
   > La valeur `/1/` de `office.js` dans l’URL du CDN préconise l’utilisation de la dernière version incrémentielle comprise dans la version 1 d’Office.js.

### <a name="update-the-manifest-file-in-your-project-to-use-schema-version-11"></a>Mise à jour du fichier manifeste dans votre projet afin d’utiliser la version 1.1 du schéma

Dans le fichier manifeste de votre complément, mettez à jour l’attribut **xmlns** de l’élément **\<OfficeApp\>** en remplaçant la valeur `1.1` de version (en conservant les attributs autres que l’attribut **xmlns inchangés** ).

```xml
<?xml version="1.0" encoding="utf-8"?>
<OfficeApp xsi:type="ContentApp"
    xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
    xmlns="http://schemas.microsoft.com/office/appforoffice/1.1">
  
  <!-- manifest contents -->

</OfficeApp>
```

> [!NOTE]
> Après avoir mis à jour la version du schéma de manifeste de complément vers la version 1.1, vous devez supprimer les éléments **Fonctionnalités** et **Fonctionnalités** et les remplacer par les éléments [Hosts](/javascript/api/manifest/hosts) et [Host](/javascript/api/manifest/host) ou les [éléments Requirements and Requirement](specify-office-hosts-and-api-requirements.md).

## <a name="updating-an-office-add-in-project-created-with-a-text-editor-or-other-ide"></a>Mise à jour d’un projet de complément Office créé avec un éditeur de texte ou une autre IDE

Pour les projets créés avant la publication de la version 1.1 de l’API JavaScript Office et du schéma manifeste du complément, vous devez mettre à jour les pages HTML de votre complément pour référencer le CDN de la bibliothèque v1.1 et mettre à jour le fichier manifeste de votre complément pour utiliser le schéma v1.1.

Le processus de mise à jour est appliqué  _par projet_  ; vous devrez répéter le processus de mise à jour pour chaque projet de complément dans lequel vous souhaitez utiliser la version 1.1 d’Office.js et du schéma de manifeste de complément.

Vous n’avez pas besoin de copies locales des fichiers d’API JavaScript Office (Office.js et des fichiers .js spécifiques à l’application) pour développer un complémentOffice (référencement du CDN pour Office.js télécharge les fichiers nécessaires au moment de l’exécution), mais si vous souhaitez obtenir une copie locale des fichiers de bibliothèque, vous pouvez utiliser [l’utilitaire Command-Line NuGet](https://docs.nuget.org/consume/installing-nuget) et la `Install-Package Microsoft.Office.js` commande pour les télécharger.

> [!NOTE]
> pour obtenir une copie du fichier XSD (définition du schéma XML) pour le manifeste de complément version 1.1, consultez les [informations de référence sur le schéma des manifestes des compléments Office (version 1.1)](../develop/add-in-manifests.md).

### <a name="update-the-office-javascript-api-library-files-in-your-project-to-use-the-newest-release"></a>Mettez à jour les fichiers de bibliothèque d’API JavaScript Office dans votre projet pour utiliser la version la plus récente

1. Ouvrez les pages HTML de votre complément dans un éditeur de texte ou une interface IDE.

2. Dans la balise **principale** des pages HTML de votre complément, commentez ou supprimez les références de script office.js existantes, puis référencez la bibliothèque d’API JavaScript Office mise à jour comme suit :

    ```html
    <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js" type="text/javascript"></script>
    ```

   > [!NOTE]
   > la valeur `/1/` devant `office.js` dans l’URL du CDN préconise l’utilisation de la dernière version incrémentielle comprise dans la version 1 d’Office.js.

### <a name="update-the-manifest-file-in-your-project-to-use-schema-version-11"></a>Mise à jour du fichier manifeste dans votre projet afin d’utiliser la version 1.1 du schéma

Dans le fichier manifeste de votre complément, mettez à jour l’attribut **xmlns** de l’élément **\<OfficeApp\>** en remplaçant la valeur `1.1` de version (en conservant les attributs autres que l’attribut **xmlns inchangés** ).

```xml
<?xml version="1.0" encoding="utf-8"?>
<OfficeApp xsi:type="ContentApp"
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
  xmlns="http://schemas.microsoft.com/office/appforoffice/1.1">
  
  <!-- manifest contents -->

</OfficeApp>
```

> [!NOTE]
> Après avoir mis à jour la version du schéma de manifeste de complément vers la version 1.1, vous devez supprimer les éléments **Fonctionnalités** et **Fonctionnalités** et les remplacer par les éléments [Hosts](/javascript/api/manifest/hosts) et [Host](/javascript/api/manifest/host) ou les [éléments Requirements and Requirement](specify-office-hosts-and-api-requirements.md).

## <a name="see-also"></a>Voir aussi

- [Spécifier les applications Office et les exigences d’API](specify-office-hosts-and-api-requirements.md) ]
- [Compréhension de l’API JavaScript pour Office](understanding-the-javascript-api-for-office.md)
- [API JavaScript pour Office](../reference/javascript-api-for-office.md)
- [Informations de référence sur le schéma des manifestes des applications pour Office (version 1.1)](../develop/add-in-manifests.md)
