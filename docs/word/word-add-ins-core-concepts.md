---
title: Modèle d’objet JavaScript Word dans les compléments Office
description: Découvrez les composants clés dans le modèle objet JavaScript spécifique à Word.
ms.date: 3/17/2022
ms.localizationpriority: high
ms.openlocfilehash: d3c2a43e2febbf31fe132dfb5c220bffcc7a1fef
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/23/2022
ms.locfileid: "63746101"
---
# <a name="word-javascript-object-model-in-office-add-ins"></a>Modèle d’objet JavaScript Word dans les compléments Office

Cet article décrit les concepts fondamentaux de l’utilisation de l’[API JavaScript Word](../reference/overview/word-add-ins-reference-overview.md) pour créer des compléments.

> [!IMPORTANT]
> Pour en savoir plus sur la nature asynchrone des API Word et la manière dont elles fonctionnent avec le document, consultez [Utilisation du modèle d’API spécifique à l’application](../develop/application-specific-api-model.md).

## <a name="officejs-apis-for-word"></a>API Office.js pour Word

Un complément Word interagit avec des objets dans Word à l’aide de l’API JavaScript Office. Cela inclut deux modèles objet JavaScript :

* **API JavaScript Word** : l’[API JavaScript Word](/javascript/api/word) fournit des objets fortement typés qui fonctionnent avec le document, les plages, les tables, les listes, la mise en forme, etc.

* **API communes** : les [API communes](/javascript/api/office)donnent accès à des fonctionnalités telles que l’interface utilisateur, les boîtes de dialogue et les paramètres client communs à plusieurs applications Office.

Vous utiliserez probablement l’API JavaScript Word pour développer la majorité des fonctionnalités des compléments destinés à Word, vous utiliserez également des objets dans l’API commune. Par exemple :

* [Office.Context](/javascript/api/office/office.context) : l’objet `Context` représente l’environnement d’exécution du complément et donne accès aux objets clés de l’API. Il se compose de détails sur la configuration du document comme `contentLanguage` et `officeTheme`, et fournit des informations sur l’environnement d’exécution du complément comme `host` et `platform`. En outre, il fournit la méthode`requirements.isSetSupported()`, que vous pouvez utiliser pour vérifier si un ensemble de conditions requises spécifié est pris en charge par l’application Word dans laquelle le complément est en cours d’exécution.
* [Office.Document](/javascript/api/office/office.document) : l’objet `Office.Document` fournit la méthode `getFileAsync()`, que vous pouvez utiliser pour télécharger le fichier Word dans lequel le complément est en cours d’exécution. Il est distinct de l’objet [Word.Document](/javascript/api/word/word.document).

![Différences entre l’API JS Word et les API courantes.](../images/word-js-api-common-api.png)

## <a name="word-specific-object-model"></a>Modèle d’objet spécifique à Word

Pour comprendre les API Word, vous devez connaître la manière dont les composants d’un document sont liés les uns aux autres.

* Le **Document** contient les **Sections**, ainsi que les entités de niveau document telles que les paramètres et les parties XML personnalisées.
* Une **Section** contient un **Corps**.
* Un **Corps** donne accès aux **Paragraphe** s, **ContentControl** s et **Plage** objets, entre autres.
* Une **Plage** représente une zone contiguë de contenu, y compris du texte, un espace vide, des **Tableaux** et des images. Elle contient également la plupart des méthodes de manipulation de texte.
* Une **Liste** représente le texte d’une liste numérotée ou une liste à puces.

## <a name="see-also"></a>Voir aussi

- [Présentation de l’API JavaScript pour Word](../reference/overview/word-add-ins-reference-overview.md)
- [Créer votre premier complément Word](../quickstarts/word-quickstart.md)
- [Didacticiel sur les compléments Word](../tutorials/word-tutorial.md)
- [Référence d’API JavaScript pour Word](/javascript/api/word)
- [Découvrez le programme pour les développeurs Microsoft 365](https://developer.microsoft.com/microsoft-365/dev-program)
