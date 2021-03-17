---
title: Rendre votre complément Office compatible avec un complément COM existant
description: Activez la compatibilité entre votre compl?ment Office et un compl?ment COM équivalent.
ms.date: 03/09/2021
localization_priority: Normal
ms.openlocfilehash: b5235255987cc6a644491bc548b92701b350a179
ms.sourcegitcommit: c0c61fe84f3c5de88bd7eac29120056bb1224fc8
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/17/2021
ms.locfileid: "50836851"
---
# <a name="make-your-office-add-in-compatible-with-an-existing-com-add-in"></a>Rendre votre complément Office compatible avec un complément COM existant

Si vous avez un compl?ment COM existant, vous pouvez créer des fonctionnalités équivalentes dans votre compl?ment Office, ce qui permet à votre solution de s’exécuter sur d’autres plateformes telles qu’Office sur le web ou Mac. Dans certains cas, il se peut que votre compl?ment Office ne soit pas en mesure de fournir toutes les fonctionnalités disponibles dans le compl?ment COM correspondant. Dans ces situations, votre compl?ment COM peut fournir une meilleure expérience utilisateur sur Windows que le compl?ment Office correspondant.

Vous pouvez configurer votre compl?ment Office de sorte que lorsque le compl?ment COM équivalent est déjà install sur l’ordinateur d’un utilisateur, Office sur Windows exécute le compl?ment COM au lieu du compl?ment Office. Le add-in COM est appelé « équivalent », car Office passe en toute transparence entre le compl?ment COM et le compl?ment Office en fonction de l’ordinateur d’un utilisateur.

> [!NOTE]
> Cette fonctionnalité est prise en charge par les plateformes suivantes, lorsqu’elle est connectée à un abonnement Microsoft 365.
>
> - Excel, Word et PowerPoint sur le web
> - Excel, Word et PowerPoint sur Windows (version 1904 ou ultérieure)
> - Excel, Word et PowerPoint sur Mac (version 13.329 ou ultérieure)
> - Outlook sur Windows (version 2102 ou ultérieure)

## <a name="specify-an-equivalent-com-add-in"></a>Spécifier un compl?ment COM équivalent

### <a name="manifest"></a>Manifeste

> [!IMPORTANT]
> S’applique à Excel, PowerPoint et Word. Prise en charge d’Outlook bientôt disponible.

Pour activer la compatibilité entre votre compl?ment Office et votre compl?ment COM, identifiez le compl?ment COM équivalent dans le manifeste de votre compl?ment Office. [](add-in-manifests.md) Ensuite, Office sur Windows utilisera le compl?ment COM au lieu du compl?ment Office, s’ils sont tous deux install s.

L’exemple suivant montre la partie du manifeste qui spécifie un compl?ment COM en tant que compl?ment équivalent. La valeur de l’élément identifie le add-in COM et l’élément `ProgId` [EquivalentAddins](../reference/manifest/equivalentaddins.md) doit être placé immédiatement avant la balise `VersionOverrides` de fermeture.

```xml
<VersionOverrides>
  ...
  <EquivalentAddins>
    <EquivalentAddin>
      <ProgId>ContosoCOMAddin</ProgId>
      <Type>COM</Type>
    </EquivalentAddin>
  </EquivalentAddins>
</VersionOverrides>
```

> [!TIP]
> Pour plus d’informations sur le module complémentaire COM et la compatibilité XLL UDF, voir Rendre vos fonctions personnalisées compatibles avec les fonctions [XLL définies par l’utilisateur.](../excel/make-custom-functions-compatible-with-xll-udf.md)

### <a name="group-policy"></a>Stratégie de groupe

> [!IMPORTANT]
> S’applique uniquement à Outlook.

Pour déclarer la compatibilité entre votre compl?ment web Outlook et le compl?ment COM/VSTO, identifiez le compl?ment COM équivalent dans la stratégie de groupe Deactiver les compl?ments web Outlook dont les compl?ments COM ou **VSTO équivalents** sont install s en configurant sur l’ordinateur de l’utilisateur. Outlook sur Windows utilisera ensuite le compl?ment COM au lieu du compl?ment web, s’ils sont tous deux install s.

1. Téléchargez le dernier [outil Modèles d’administration,](https://www.microsoft.com/download/details.aspx?id=49030)en vous important des instructions d’installation **de l’outil.**
1. Ouvrez l’Éditeur de stratégie de groupe local (**gpedit.msc**).
1. Accédez **à Modèles** d’administration de configuration  >     >  **utilisateur Microsoft Outlook 2016**  >  **Divers.**
1. Sélectionnez le paramètre Désactiver les **compl?ments web Outlook** dont les compl?ments COM ou VSTO équivalents sont install s .
1. Ouvrez le lien pour modifier le paramètre de stratégie.
1. Dans la boîte **de dialogue, les applications web Outlook** sont à désactiver :
    1. Définissez **le nom de** la valeur sur la valeur trouvée dans le manifeste du `Id` add-in web. **Important**: *n’ajoutez* pas d’accolades `{}` autour de l’entrée.
    1. Définissez **la** valeur sur la valeur du `ProgId` compl?ment COM/VSTO équivalent.
    1. Sélectionnez **OK** pour mettre la mise à jour en vigueur.
    ![Capture d’écran montrant la boîte de dialogue « Les applications web Outlook à désactiver »](../images/outlook-deactivate-gpo-dialog.png)

## <a name="equivalent-behavior-for-users"></a>Comportement équivalent pour les utilisateurs

Lorsqu’un compl?ment [COM](#specify-an-equivalent-com-add-in)équivalent est spécifié, Office sur Windows n’affiche pas l’interface utilisateur de votre compl?ment Office si le compl?ment COM équivalent est install . Office masque uniquement les boutons du ruban du add-in Office et n’empêche pas l’installation. Par conséquent, votre add-in Office apparaîtra toujours aux emplacements suivants dans l’interface utilisateur :

- Sous **Mes modules**
- En tant qu’entrée dans le gestionnaire du ruban (Excel, Word et PowerPoint uniquement)

> [!NOTE]
> La spécification d’un équivalent com dans le manifeste n’a aucun effet sur les autres plateformes telles qu’Office sur le web ou sur Mac.

Les scénarios suivants décrivent ce qui se produit en fonction de la façon dont l’utilisateur acquiert le add-in Office.

### <a name="appsource-acquisition-of-an-office-add-in"></a>Acquisition d’un add-in Office dans AppSource

Si un utilisateur acquiert le compl?ment Office auprès d’AppSource et que le compl?ment COM équivalent est déjà install ? , Office :

1. Installez le add-in Office.
2. Masquer l’interface utilisateur du add-in Office dans le ruban.
3. Affichez un appel pour l’utilisateur qui pointe sur le bouton du ruban du compl?ment COM.

### <a name="centralized-deployment-of-office-add-in"></a>Déploiement centralisé d’un add-in Office

Si un administrateur déploie le add-in Office sur son client à l’aide d’un déploiement centralisé et que le module com équivalent est déjà installé, l’utilisateur doit redémarrer Office avant de voir des modifications. Après le redémarrage d’Office, il :

1. Installez le add-in Office.
2. Masquer l’interface utilisateur du add-in Office dans le ruban.
3. Affichez un appel pour l’utilisateur qui pointe sur le bouton du ruban du compl?ment COM.

### <a name="document-shared-with-embedded-office-add-in"></a>Document partagé avec un add-in Office incorporé

Si un utilisateur a installé le compl?ment COM, puis obtient un document partagé avec le compl?ment Office incorporé, alors lorsqu’il ouvre le document, Office :

1. Invitez l’utilisateur à faire confiance au add-in Office.
2. S’il est approuvé, le add-in Office s’installe.
3. Masquer l’interface utilisateur du add-in Office dans le ruban.

## <a name="other-com-add-in-behavior"></a>Comportement des autres compl?ments COM

### <a name="excel-powerpoint-word"></a>Excel, PowerPoint, Word

Si un utilisateur désinstalle le compl?ment COM équivalent, Office sur Windows restaure l’interface utilisateur du compl?ment Office.

Après avoir spécifié un compl?ment COM équivalent pour votre compl?ment Office, Office cesse de traiter les mises à jour pour votre compl?ment Office. Pour obtenir les dernières mises à jour pour le compl?ment Office, l’utilisateur doit d’abord désinstaller le compl?ment COM.

### <a name="outlook"></a>Outlook

Le add-in COM/VSTO doit être connecté au moment du début d’Outlook afin que le compl?ment web correspondant soit désactivé.

Si le compl?ment COM/VSTO est alors déconnecté lors d’une session Outlook suivante, le compl?ment web restera probablement désactivé jusqu’au redémarrage d’Outlook.

## <a name="see-also"></a>Voir aussi

- [Rendre vos fonctions personnalisées compatibles avec les fonctions XLL définies par l’utilisateur](../excel/make-custom-functions-compatible-with-xll-udf.md)
