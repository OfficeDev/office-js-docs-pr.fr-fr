---
title: Rendre votre complément Office compatible avec un complément COM existant
description: Activer la compatibilité entre votre complément Office et un complément COM équivalent
ms.date: 06/19/2019
localization_priority: Normal
ms.openlocfilehash: a18adb9841a9580d77c5110a0346f365e38e3746
ms.sourcegitcommit: 4bf5159a3821f4277c07d89e88808c4c3a25ff81
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/18/2019
ms.locfileid: "35059719"
---
# <a name="make-your-office-add-in-compatible-with-an-existing-com-add-in-preview"></a>Faire en sorte que votre complément Office soit compatible avec un complément COM existant (aperçu)

Si vous disposez d’un complément COM existant, vous pouvez créer une fonctionnalité équivalente dans votre complément Office, ce qui permet à votre solution de s’exécuter sur d’autres plateformes, telles qu’Office sur le Web ou Office sur Mac. Dans certains cas, votre complément Office peut ne pas être en mesure de fournir toutes les fonctionnalités disponibles dans le complément COM correspondant. Dans ce cas, votre complément COM peut fournir une meilleure expérience utilisateur sur Windows que le complément Office correspondant.

Vous pouvez configurer votre complément Office de sorte que, lorsque le complément COM équivalent est déjà installé sur l’ordinateur d’un utilisateur, Office sur Windows exécute le complément COM au lieu du complément Office. Le complément COM est appelé «équivalent», car Office effectuera une transition transparente entre le complément COM et le complément Office en fonction de celui sur lequel est installé l’ordinateur d’un utilisateur.

> [!NOTE]
> Cette fonctionnalité est actuellement en préversion et n’est pas prise en charge dans les environnements de production. Elle est disponible dans Excel, Word et PowerPoint version 16.0.11629.20214 ou ultérieure. Pour accéder à cette version, vous devez disposer d’un abonnement Office 365 et rejoindre le programme [Office](https://products.office.com/office-insider) Insider au niveau Insider. ****

## <a name="specify-an-equivalent-com-add-in-in-the-manifest"></a>Spécifier un complément COM équivalent dans le manifeste

Pour activer la compatibilité entre votre complément Office et le complément COM, identifiez le complément COM équivalent dans le [manifeste](add-in-manifests.md) de votre complément Office. Office sur Windows utilisera ensuite le complément COM au lieu du complément Office, s’ils sont tous les deux installés.

L’exemple suivant montre la partie du manifeste qui spécifie un complément COM sous la forme d’un complément équivalent. La valeur de l' `ProgId` élément identifie le complément COM et l' `EquivalentAddins` élément doit être placé immédiatement avant la balise de `VersionOverrides` fermeture.

```xml
<VersionOverrides>
  ...
  <EquivalentAddins>
    <EquivalentAddin>
      <ProgId>ContosoCOMAddin</ProgId>
      <Type>COM</Type>
    </EquivalentAddin>
  <EquivalentAddins>
</VersionOverrides>
```

> [!TIP]
> Pour plus d’informations sur les compléments COM et la compatibilité des FDU XLL, consultez [la rubrique faire en sorte que les fonctions personnalisées soient compatibles avec les fonctions XLL définies par l’utilisateur](../excel/make-custom-functions-compatible-with-xll-udf.md).

## <a name="equivalent-behavior-for-users"></a>Comportement équivalent pour les utilisateurs

Lorsqu’un complément COM équivalent est spécifié dans le manifeste du complément Office, Office sur Windows n’affiche pas l’interface utilisateur (IU) de votre complément Office si le complément COM équivalent est installé. Office masque uniquement les boutons du ruban du complément Office et n’empêche pas l’installation. Par conséquent, votre complément Office continuera à apparaître aux emplacements suivants au sein de l’interface utilisateur:

- Sous **mes compléments**
- Comme entrée dans le gestionnaire de ruban

> [!NOTE]
> La spécification d’un complément COM équivalent dans le manifeste n’a aucun effet sur les autres plateformes comme Office sur le Web ou Office pour Mac.

Les scénarios suivants décrivent ce qui se produit en fonction de la manière dont l’utilisateur acquiert le complément Office.

### <a name="appsource-acquisition-of-an-office-add-in"></a>AppSource acquisition d’un complément Office

Si un utilisateur acquiert le complément Office à partir de AppSource et que le complément COM équivalent est déjà installé, Office:

1. Installez le complément Office.
2. Masquer l’interface utilisateur du complément Office dans le ruban.
3. Afficher un appel pour l’utilisateur qui pointe vers le bouton du ruban de complément COM.

### <a name="centralized-deployment-of-office-add-in"></a>Déploiement centralisé du complément Office

Si un administrateur déploie le complément Office sur son client à l’aide d’un déploiement centralisé, et que le complément COM équivalent est déjà installé, l’utilisateur doit redémarrer Office avant de voir les modifications. Après le redémarrage d’Office, il peut:

1. Installez le complément Office.
2. Masquer l’interface utilisateur du complément Office dans le ruban.
3. Afficher un appel pour l’utilisateur qui pointe vers le bouton du ruban de complément COM.

### <a name="document-shared-with-embedded-office-add-in"></a>Document partagé avec un complément Office incorporé

Si un utilisateur a installé le complément COM, puis qu’il obtient un document partagé avec le complément Office incorporé, lorsqu’il ouvre le document, Office:

1. Inviter l’utilisateur à approuver le complément Office.
2. S’il est approuvé, le complément Office est installé.
3. Masquer l’interface utilisateur du complément Office dans le ruban.

## <a name="other-com-add-in-behavior"></a>Autre comportement de complément COM

Si un utilisateur désinstalle le complément COM équivalent, Office sur Windows restaure l’interface utilisateur du complément Office.

Une fois que vous avez spécifié un complément COM équivalent pour votre complément Office, Office cesse de traiter les mises à jour pour votre complément Office. Pour obtenir les dernières mises à jour pour le complément Office, l’utilisateur doit d’abord désinstaller le complément COM.

## <a name="see-also"></a>Voir aussi

- [Faire en sorte que vos fonctions personnalisées soient compatibles avec les fonctions XLL définies par l’utilisateur](../excel/make-custom-functions-compatible-with-xll-udf.md)
