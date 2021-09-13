---
title: Valider un manifeste de complément Office
description: Découvrez comment valider le manifeste d’un Office à l’aide du schéma XML et d’autres outils.
ms.date: 09/18/2020
ms.localizationpriority: medium
ms.openlocfilehash: 60151522906fc45c8d0d501c8263acd818612897
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/12/2021
ms.locfileid: "59152232"
---
# <a name="validate-an-office-add-ins-manifest"></a>Valider un manifeste de complément Office

Vous souhaitez peut-être valider le fichier manifeste de votre complément pour vous assurer que celui-ci est correct et complet. La validation peut également identifier les problèmes à l’origine de l’erreur « Votre manifeste de complément n’est pas valide » lorsque vous essayez de charger une version test de votre complément. Cet article décrit plusieurs méthodes de validation du fichier manifeste.

> [!NOTE]
> Pour en savoir plus sur l’utilisation de la journalisation de l’exécution pour résoudre des problèmes relatifs au manifeste de votre complément, consultez [Déboguer votre complément avec la journalisation de l’exécution](runtime-logging.md).

## <a name="validate-your-manifest-with-the-yeoman-generator-for-office-add-ins"></a>Valider votre manifeste avec le générateur Yeoman pour les compléments Office

Si vous avez utilisé [le générateur Yeoman pour les compléments Office](https://www.npmjs.com/package/generator-office) pour créer votre complément, vous pouvez également l’utiliser pour valider le fichier manifeste de votre projet. Exécutez la commande suivante dans le répertoire racine de votre projet.

```command&nbsp;line
npm run validate
```

![Image GIF animée qui montre le validateur Office Yo exécuté sur la ligne de commande et générant des résultats qui indiquent validation transmise.](../images/yo-office-validator.gif)

> [!NOTE]
> Pour accéder à cette fonctionnalité, votre projet de complément doit être créé à l’aide du [générateur Yeoman pour les compléments Office](https://www.npmjs.com/package/generator-office) (version 1.1.17 ou ultérieure).

## <a name="validate-your-manifest-with-office-addin-manifest"></a>Valider votre manifeste avec office-addin-manifest

Si vous n’avez pas utilisé [le générateur Yeoman pour les compléments Office](https://www.npmjs.com/package/generator-office) pour créer votre complément, vous pouvez valider le fichier manifeste à l’aide de [office-addin-manifest](https://www.npmjs.com/package/office-addin-manifest).

1. Installez [Node.js](https://nodejs.org/download/).

1. Ouvrez une invite de commandes et installez le validateur avec la commande suivante.

    ```command&nbsp;line
    npm install -g office-addin-manifest
    ```

1. Exécutez la commande suivante *dans le répertoire racine de votre projet.*

    ```command&nbsp;line
    npm run validate
    ```

    > [!NOTE]
    > Si cette commande n’est pas disponible ou ne fonctionne pas, exécutez la commande suivante à la place pour forcer l’utilisation de la dernière version de l’outil de manifeste du addin office (en remplaçant par le nom du fichier `MANIFEST_FILE` manifeste).
    >
    > ```command&nbsp;line
    > npx --ignore-existing office-addin-manifest validate MANIFEST_FILE
    > ```

## <a name="validate-your-manifest-against-the-xml-schema"></a>Validez votre manifeste par rapport au schéma XML

Vous pouvez valider le fichier manifeste par rapport aux fichiers de [définition de schéma XML (XSD)](/openspecs/office_file_formats/ms-owemxml/c6a06390-34b8-4b42-82eb-b28be12494a8). Cela permet de s’assurer que le fichier manifeste suit le schéma approprié, y compris les espaces de noms pour les éléments que vous utilisez. Si vous avez copié des éléments à partir d’autres exemples de manifestes, vérifiez par deux fois que vous avez également **inclus les espaces de noms appropriés**. Pour ce faire, vous pouvez utiliser un outil de validation de schéma XML.

### <a name="to-use-a-command-line-xml-schema-validation-tool-to-validate-your-manifest"></a>Pour utiliser un outil de validation de schéma XML à ligne de commande pour valider votre manifeste

1. Installez [tar](https://www.gnu.org/software/tar/) et [libxml](http://xmlsoft.org/FAQ.html), si vous ne l’avez pas déjà fait.

1. Exécutez la commande suivante. Remplacez `XSD_FILE` par le chemin d’accès au fichier XSD manifeste et `XML_FILE` par le chemin d’accès au fichier XML manifeste.

    ```command&nbsp;line
    xmllint --noout --schema XSD_FILE XML_FILE
    ```

## <a name="see-also"></a>Voir aussi

- [Manifeste XML des compléments Office](../develop/add-in-manifests.md)
- [Vider le cache Office](clear-cache.md)
- [Déboguer votre complément avec la journalisation de l’exécution](runtime-logging.md)
- [Chargement de la version test des compléments Office](sideload-office-add-ins-for-testing.md)
- [Débogage des compléments Office](debug-add-ins-using-f12-developer-tools-on-windows-10.md)
