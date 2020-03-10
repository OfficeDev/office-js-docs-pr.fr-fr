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
# <a name="validate-an-office-add-ins-manifest"></a>Valider un manifeste de complément Office

Vous souhaitez peut-être valider le fichier manifeste de votre complément pour vous assurer que celui-ci est correct et complet. La validation peut également identifier les problèmes à l’origine de l’erreur « Votre manifeste de complément n’est pas valide » lorsque vous essayez de charger une version test de votre complément. Cet article décrit plusieurs méthodes de validation du fichier manifeste.

> [!NOTE]
> Pour en savoir plus sur l’utilisation de la journalisation de l’exécution pour résoudre des problèmes relatifs au manifeste de votre complément, consultez [Déboguer votre complément avec la journalisation de l’exécution](runtime-logging.md).

## <a name="validate-your-manifest-with-the-yeoman-generator-for-office-add-ins"></a>Valider votre manifeste avec le générateur Yeoman pour les compléments Office

Si vous avez utilisé [le générateur Yeoman pour les compléments Office](https://www.npmjs.com/package/generator-office) pour créer votre complément, vous pouvez également l’utiliser pour valider le fichier manifeste de votre projet. Exécutez la commande suivante dans le répertoire racine de votre projet :

```command&nbsp;line
npm run validate
```

![Gif animé qui montre le validateur Yo Office exécuté sur la ligne de commande et les résultats générés indiquant « Validation Passed » (validation réussie)](../images/yo-office-validator.gif)

> [!NOTE]
> Pour accéder à cette fonctionnalité, votre projet de complément doit être créé à l’aide du [générateur Yeoman pour les compléments Office](https://www.npmjs.com/package/generator-office) (version 1.1.17 ou ultérieure).

## <a name="validate-your-manifest-with-office-addin-manifest"></a>Valider votre manifeste avec office-addin-manifest

Si vous n’avez pas utilisé [le générateur Yeoman pour les compléments Office](https://www.npmjs.com/package/generator-office) pour créer votre complément, vous pouvez valider le fichier manifeste à l’aide de [office-addin-manifest](https://www.npmjs.com/package/office-addin-manifest).

1. Installez [Node.js](https://nodejs.org/download/).

2. Exécutez la commande suivante dans le répertoire racine de votre projet. Remplacez `MANIFEST_FILE` par le nom du fichier manifeste.

    ```command&nbsp;line
    npx office-addin-manifest validate MANIFEST_FILE
    ```

    > [!NOTE]
    > Si elle s’exécute, la commande renvoie le message d’erreur « La syntaxe de la commande n’est pas valide » (étant donné que la commande `validate` n’est pas reconnue), exécutez la commande suivante pour valider le manifeste (en remplaçant `MANIFEST_FILE` par le nom du fichier manifeste) : 
    >
    > `npx --ignore-existing office-addin-manifest validate MANIFEST_FILE`

## <a name="validate-your-manifest-against-the-xml-schema"></a>Validez votre manifeste par rapport au schéma XML

Vous pouvez valider le fichier manifeste par rapport aux fichiers de [définition de schéma XML (XSD)](/openspecs/office_file_formats/ms-owemxml/c6a06390-34b8-4b42-82eb-b28be12494a8). Cela permet de s’assurer que le fichier manifeste suit le schéma approprié, y compris les espaces de noms pour les éléments que vous utilisez. Si vous avez copié des éléments à partir d’autres exemples de manifestes, vérifiez par deux fois que vous avez également **inclus les espaces de noms appropriés**. Pour ce faire, vous pouvez utiliser un outil de validation de schéma XML.

### <a name="to-use-a-command-line-xml-schema-validation-tool-to-validate-your-manifest"></a>Pour utiliser un outil de validation de schéma XML à ligne de commande pour valider votre manifeste

1. Installez [tar](https://www.gnu.org/software/tar/) et [libxml](http://xmlsoft.org/FAQ.html), si vous ne l’avez pas déjà fait.

2. Exécutez la commande suivante. Remplacez `XSD_FILE` par le chemin d’accès au fichier XSD manifeste et `XML_FILE` par le chemin d’accès au fichier XML manifeste.
    
    ```command&nbsp;line
    xmllint --noout --schema XSD_FILE XML_FILE
    ```

## <a name="see-also"></a>Voir aussi

- [Manifeste XML des compléments Office](../develop/add-in-manifests.md)
- [Vider le cache Office](clear-cache.md)
- [Déboguer votre complément avec la journalisation de l’exécution](runtime-logging.md)
- [Chargement de la version test des compléments Office](sideload-office-add-ins-for-testing.md)
- [Débogage des compléments Office](debug-add-ins-using-f12-developer-tools-on-windows-10.md)