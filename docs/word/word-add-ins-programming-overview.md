---
title: Présentation des compléments Word
description: Découvrez les concepts de base des compléments Word.
ms.date: 10/14/2020
ms.topic: conceptual
ms.custom: scenarios:getting-started
ms.localizationpriority: high
ms.openlocfilehash: 1f55977cba42c1c16a8533958f60b6da0e9a3650
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/12/2021
ms.locfileid: "59150283"
---
# <a name="word-add-ins-overview"></a>Présentation des compléments Word

Vous souhaitez créer une solution qui étend les fonctionnalités de Word ? Par exemple, une solution qui assemble automatiquement les documents ? Ou une solution qui relie les données et y accède dans un document Word à partir d’autres sources de données ? Vous pouvez utiliser la plateforme de compléments Office. Elle comprend l’API JavaScript pour Word et l’API Office JavaScript, pour développer les clients Word qui s’exécutent sur un ordinateur de bureau Windows, un Mac ou dans le cloud.

Les compléments Word font partie des nombreuses options de développement disponibles sur la [plateforme de compléments Office](../overview/office-add-ins.md). Vous pouvez utiliser les commandes de complément pour développer l’interface utilisateur Word et créer des volets Office qui exécutent un code JavaScript pour interagir avec le contenu d’un document Word. Tout code que vous pouvez exécuter dans un navigateur peut s’exécuter dans un complément Word. Les compléments qui interagissent avec le contenu d’un document Word créent des requêtes qui agissent sur des objets Word et synchronisent l’état des objets.

[!INCLUDE [publish policies note](../includes/note-publish-policies.md)]

La figure suivante montre un exemple d’un complément Word qui s’exécute dans un volet des tâches.

*Figure 1. Complément exécuté dans un volet Office de Word*

![Complément exécuté dans un volet Office de Word.](../images/word-add-in-show-host-client.png)

Le complément Word (1) peut envoyer des demandes dans le document Word (2) et utiliser JavaScript pour accéder à l’objet de paragraphe et mettre à jour, supprimer ou déplacer le paragraphe. Par exemple, le code suivant montre comment ajouter une nouvelle phrase à ce paragraphe.

```js
Word.run(function (context) {
    var paragraphs = context.document.getSelection().paragraphs;
    paragraphs.load();
    return context.sync().then(function () {
        paragraphs.items[0].insertText(' New sentence in the paragraph.',
                                       Word.InsertLocation.end);
    }).then(context.sync);
});

```

Vous pouvez utiliser n’importe quelle technologie de serveur web pour héberger votre complément Word, comme ASP.NET, NodeJS ou Python. Utilisez votre infrastructure côté client préférée (Ember, Backbone, Angular, React), ou utilisez VanillaJS pour développer votre solution et utilisez des services comme Azure pour [authentifier](../develop/overview-authn-authz.md) et héberger votre application.

Les interfaces API JavaScript pour Word permettent à votre application d’accéder aux objets et aux métadonnées situés dans le document Word. Vous pouvez utiliser ces API pour créer des compléments destinés à :

* Word 2013 ou version ultérieure sur Windows
* Word sur le web
* Word 2016 ou version ultérieure sur Mac
* Word sur iPad

Écrivez votre complément une seule fois. Celui-ci s’exécutera dans toutes les versions de Word sur plusieurs plateformes. Pour plus d’informations, voir [Disponibilité des compléments Office sur les plateformes et les applications clientes](../overview/office-add-in-availability.md).

## <a name="javascript-apis-for-word"></a>APIs JavaScript pour Word

Vous pouvez utiliser deux ensembles d’API JavaScript pour interagir avec les objets et les métadonnées dans un document Word. La première est l’[API commune](/javascript/api/office), qui a été introduite dans Office 2013. La plupart des objets de l’API Commune peuvent être utilisés dans des compléments hébergés par au moins deux clients Office. Cette API utilise beaucoup les rappels.

Le deuxième est l’[API JavaScript pour Word](/javascript/api/word) qui est un [modèle d’API spécifique à l’application](../develop/application-specific-api-model.md) introduit avec Word 2016. Il s’agit d’un modèle objet fortement typé qui vous permet de créer des compléments Word destinés à Word 2016 sur Mac et Windows. Ce modèle objet utilise les promesses et fournit un accès aux objets Word, tels que le [corps](/javascript/api/word/word.body), les [contrôles de contenu](/javascript/api/word/word.contentcontrol), les [images incorporées](/javascript/api/word/word.inlinepicture) et les [paragraphes](/javascript/api/word/word.paragraph). L’API JavaScript pour Word inclut des définitions TypeScript et des fichiers vsdoc pour vous permettre d’obtenir des conseils concernant votre code dans votre environnement de développement intégré (IDE).

Actuellement, tous les clients Word prennent en charge l’API JavaScript Office partagée, et la plupart des clients prennent en charge l’API JavaScript Word. Pour plus d’informations sur les clients pris en charge, consultez [Disponibilité de l’application cliente Office et de la plateforme pour les compléments Office](../overview/office-add-in-availability.md).

Nous vous recommandons de démarrer avec l’API JavaScript pour Word car le modèle d’objet est plus facile à utiliser. Utilisez l’API JavaScript pour Word pour :

* Accéder aux objets d’un document Word.

Utilisez l’API Office JavaScript partagée pour :

* Cibler Word 2013.
* Effectuer des actions initiales pour l’application.
* Vérifier l’ensemble de conditions requises pris en charge.
* Accéder aux métadonnées, aux paramètres et aux informations de l’environnement du document.
* Établir des liaisons avec des sections d’un document et capturer les événements.
* Utiliser des parties XML personnalisées.
* Ouvrir une boîte de dialogue.

## <a name="next-steps"></a>Étapes suivantes

Prêt à créer votre premier complément Word ? Consultez la page [Création de votre premier complément Word](../quickstarts/word-quickstart.md). Utilisez le [manifeste de complément](../develop/add-in-manifests.md) pour décrire l’emplacement d’hébergement de votre complément et son affichage, et définir des autorisations et d’autres informations.

Pour savoir comment concevoir un complément Word de qualité qui offre une expérience intéressante aux utilisateurs, consultez les [recommandations de conception](../design/add-in-design.md) et les [meilleures pratiques](../concepts/add-in-development-best-practices.md).

Une fois le développement de votre complément terminé, vous pouvez le [publier](../publish/publish.md) sur un partage réseau, dans un catalogue d’applications ou dans AppSource.

## <a name="see-also"></a>Voir aussi

* [Développement de compléments Office](../develop/develop-overview.md)
* [Découvrez le programme pour les développeurs Microsoft 365](https://developer.microsoft.com/microsoft-365/dev-program)
* [Vue d’ensemble de la plateforme des compléments Office](../overview/office-add-ins.md)
* [Référence d’API JavaScript pour Word](../reference/overview/word-add-ins-reference-overview.md)