---
title: Faire en sorte que votre complément Office soit compatible avec un complément COM existant
description: Activer la compatibilité avec un complément COM équivalent doté de la même fonctionnalité que votre complément Office
ms.date: 04/22/2019
localization_priority: Normal
ms.openlocfilehash: 8f3780814163cc4dd21311b362d1d821a14b3e80
ms.sourcegitcommit: 7462409209264dc7f8f89f3808a7a6249fcd739e
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/26/2019
ms.locfileid: "33356863"
---
# <a name="make-your-office-add-in-compatible-with-an-existing-com-add-in"></a>Faire en sorte que votre complément Office soit compatible avec un complément COM existant

Si vous disposez d'un complément COM existant, vous pouvez créer une fonctionnalité équivalente dans votre complément Office pour étendre les fonctionnalités de votre solution à d'autres plateformes, comme Online ou macOS. Toutefois, les compléments Office ne disposent pas de toutes les fonctionnalités disponibles dans les compléments COM. Votre complément COM peut fournir une meilleure expérience que le complément Office sur Windows dans Excel, Word et PowerPoint.

Vous pouvez configurer votre complément Office de sorte que, lorsqu'un complément COM équivalent est déjà installé sur l'ordinateur de l'utilisateur, Office exécute le complément COM au lieu de votre complément Office. Le complément COM est appelé «équivalent», car Office effectuera une transition transparente entre le complément COM et le complément Office en fonction de ce qui est installé sur Windows.

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="specify-an-equivalent-com-add-in-in-the-manifest"></a>Spécifier un complément COM équivalent dans le manifeste

Pour activer la compatibilité avec un complément COM existant, identifiez le complément COM équivalent dans le manifeste de votre complément Office. Office utilise ensuite le complément COM au lieu de votre complément Office lors de l'exécution de Windows.

Spécifiez `ProgID` le du complément COM équivalent. Office utilise alors l'interface utilisateur du complément COM au lieu de l'interface utilisateur de votre complément Office lorsque le complément COM est installé.

L'exemple suivant montre comment spécifier un complément COM et un XLL comme équivalent. Souvent, vous spécifierez à la fois de manière à ce que cet exemple montre les deux dans le contexte. Ils sont identifiés par leur `ProgID` et `FileName` respectivement. Pour plus d'informations sur la compatibilité des XLL, consultez [la rubrique faire en sorte que vos fonctions personnalisées soient compatibles avec les fonctions XLL définies par l'utilisateur](../excel/make-custom-functions-compatible-with-xll-udf.md).

```xml
<VersionOverrides>
...
<EquivalentAddins>
  <EquivalentAddin>
    <ProgID>ContosoCOMAddin</ProgID>
    <Type>COM</Type>
  </EquivalentAddin>

  <EquivalentAddin>
    <FileName>contosofunctions.xll</FileName>
    <Type>XLL</Type>
  </EquivalentAddin>
<EquivalentAddins>
...
</VersionOverrides>
```

## <a name="equivalent-behavior-for-users"></a>Comportement équivalent pour les utilisateurs

Lorsqu'un complément COM équivalent est spécifié dans le manifeste du complément Office, Office supprime l'interface utilisateur de votre complément Office sur Windows lorsque le complément COM équivalent est installé. Cela n'affecte pas l'interface utilisateur de votre complément Office sur d'autres plateformes, comme Online ou macOS. Office masque uniquement les boutons du ruban et n'empêche pas l'installation. Par conséquent, votre complément Office continuera à apparaître aux emplacements d'IU suivants:

- Sous **My Add-ins** car il est techniquement installé.
- Comme entrée dans le gestionnaire de ruban.

Les scénarios suivants décrivent ce qui se produit en fonction de la manière dont l'utilisateur acquiert le complément Office.

### <a name="appsource-acquisition-of-an-office-add-in"></a>AppSource acquisition d'un complément Office

Si un utilisateur télécharge le complément Office à partir de AppSource, et que le complément COM équivalent est déjà installé, Office:

1. Installez le complément Office.
2. Masquer l'interface utilisateur du complément Office dans le ruban.
3. Afficher un appel pour l'utilisateur qui pointe vers le bouton du ruban de complément COM.

### <a name="centralized-deployment-of-office-add-in"></a>Déploiement centralisé du complément Office

Si un administrateur déploie le complément Office sur son client à l'aide d'un déploiement centralisé, et que le complément COM équivalent est déjà installé, l'utilisateur doit redémarrer Office pour qu'il voit les modifications. Après le redémarrage d'Office, il peut:

1. Installez le complément Office.
2. Masquer l'interface utilisateur du complément Office dans le ruban.
3. Afficher un appel pour l'utilisateur qui pointe vers le bouton du ruban de complément COM.

### <a name="document-shared-with-embedded-office-add-in"></a>Document partagé avec un complément Office incorporé

Si un utilisateur a installé le complément COM, puis qu'il obtient un document partagé avec le complément Office incorporé, lorsqu'il ouvre le document, Office:

1. Inviter l'utilisateur à approuver le complément Office.
2. S'il est approuvé, le complément Office est installé.
3. Masquer l'interface utilisateur du complément Office dans le ruban.

## <a name="other-com-add-in-behavior"></a>Autre comportement de complément COM

Si un utilisateur désinstalle le complément COM, office restaure l'interface utilisateur d'un complément Office sur Windows pour le complément Office installé équivalente.

Une fois que vous avez spécifié un complément COM équivalent pour votre complément Office, Office cesse de traiter les mises à jour pour votre complément Office. L'utilisateur doit désinstaller l'ordre des compléments COM pour obtenir les dernières mises à jour pour le complément Office.

## <a name="see-also"></a>Voir aussi

- [Faire en sorte que vos fonctions personnalisées soient compatibles avec les fonctions XLL définies par l'utilisateur](../excel/make-custom-functions-compatible-with-xll-udf.md)
