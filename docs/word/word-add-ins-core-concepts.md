---
title: Modèle d’objet JavaScript Word dans les compléments Office
description: Découvrez les classes les plus importantes dans le modèle objet JavaScript spécifique à Word.
ms.date: 10/14/2020
localization_priority: Priority
ms.openlocfilehash: c85c56987ef5de7c087064ac668f137326089642
ms.sourcegitcommit: 42e6cfe51d99d4f3f05a3245829d764b28c46bbb
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/23/2020
ms.locfileid: "48740867"
---
# <a name="word-javascript-object-model-in-office-add-ins"></a>Modèle d’objet JavaScript Word dans les compléments Office

Cet article décrit les concepts de base de l’utilisation de [l’API JavaScript pour Word](../reference/overview/word-add-ins-reference-overview.md) pour créer des compléments. Il présente les concepts fondamentaux de l’utilisation de l’API.

> [!IMPORTANT]
> Pour en savoir plus sur la nature asynchrone des API Word et la manière dont elles fonctionnent avec le document, consultez [Utilisation du modèle d’API spécifique à l’application](../develop/application-specific-api-model.md).

## <a name="officejs-apis-for-word"></a>API Office.js pour Word

Un complément Word interagit avec des objets dans Excel en utilisant l’API Office JavaScript, qui inclut deux modèles d’objets JavaScript :

* **API JavaScript Word** : l’[API JavaScript Word](../reference/overview/word-add-ins-reference-overview.md) fournit des objets fortement typés que vous pouvez utiliser pour accéder au document, à des plages, à des tableaux, à des listes, à une mise en forme, etc.

* **API communes** : l’[API commune](/javascript/api/office) peut être utilisée pour accéder à des fonctionnalités telles que l’interface utilisateur, les boîtes de dialogue et les paramètres de client communs à différents types d’applications Office.

Vous utiliserez probablement l’API JavaScript Word pour développer la majorité des fonctionnalités des compléments destinés à Word, vous utiliserez également des objets dans l’API commune. Par exemple :

* [Context](/javascript/api/office/office.context) :le `Context` représente l’environnement d’exécution du complément et permet d’accéder à des objets clés de l’API. Il se compose de détails sur la configuration du document comme `contentLanguage` et `officeTheme`, et fournit des informations sur l’environnement d’exécution du complément comme `host` et `platform`. En outre, il fournit la méthode `requirements.isSetSupported()` que vous pouvez utiliser pour vérifier si un ensemble de conditions requises spécifié est pris en charge par l’application Excel dans laquelle le complément est exécuté.
* [Document](/javascript/api/office/office.document) : le `Document` fournit la méthode `getFileAsync()` que vous pouvez utiliser pour télécharger le fichier Word dans lequel le complément est exécuté.

![Image des différences entre l’API Word et les API communes](../images/word-js-api-common-api.png)

## <a name="word-specific-object-model"></a>Modèle d’objet spécifique à Word

Pour comprendre les API Word, vous devez connaître la manière dont les composants d’un document sont liés les uns aux autres.

* Le **Document** contient les **Sections**, ainsi que les entités de niveau document telles que les paramètres et les parties XML personnalisées.
* Une **Section** contient un **Corps**.
* Un **Corps** donne accès aux **Paragraphe**s, **ContentControl**s et **Plage** objets, entre autres.
* Une **Plage** représente une zone contiguë de contenu, y compris du texte, un espace vide, des **Tableaux** et des images. Elle contient également la plupart des méthodes de manipulation de texte.
* Une **Liste** représente le texte d’une liste numérotée ou une liste à puces.

## <a name="see-also"></a>Voir aussi

- [Présentation de l’API JavaScript pour Word](../reference/overview/word-add-ins-reference-overview.md)
- [Créer votre premier complément Word](../quickstarts/word-quickstart.md)
- [Didacticiel sur les compléments Word](../tutorials/word-tutorial.md)
- [Référence sur l’API JavaScript pour Word](/javascript/api/word)
- [Découvrez le programme pour les développeurs Microsoft 365](https://developer.microsoft.com/microsoft-365/dev-program)