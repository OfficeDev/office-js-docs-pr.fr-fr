---
title: Élément IconUrl dans le fichier manifeste
description: L’élément IconUrl spécifie l’URL de l’image qui représente votre complément Office dans l’expérience utilisateur d’insertion et dans l’Office Store.
ms.date: 06/20/2019
localization_priority: Normal
ms.openlocfilehash: 2ccfc2fc1d0a07f6d549f388bbb58e40e79a17d5
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/08/2020
ms.locfileid: "44611791"
---
# <a name="iconurl-element"></a>IconUrl, élément

Spécifie l’URL de l’image utilisée pour représenter votre complément Office dans l’UX d’insertion UX et l’Office Store.

**Type de complément :** application de contenu, de volet Office, de messagerie

## <a name="syntax"></a>Syntaxe

```XML
<IconUrl DefaultValue="string" />
```

## <a name="can-contain"></a>Peut contenir

[Override](override.md)

## <a name="attributes"></a>Attributs

|**Attribut**|**Type**|**Obligatoire**|**Description**|
|:-----|:-----|:-----|:-----|
|DefaultValue|chaîne|obligatoire|Spécifie la valeur par défaut de ce paramètre, exprimée pour les paramètres régionaux spécifiés dans l’élément [DefaultLocale](defaultlocale.md).|

## <a name="remarks"></a>Remarques

Pour un complément de messagerie **, l’icône**est affichée dans l’interface utilisateur de gestion des  >  **compléments** (Outlook) ou **paramètres**  >  **gérer les compléments** (Outlook sur le Web). Pour un complément de contenu ou de volet Office, l’icône s’affiche dans l’interface utilisateur, sous **Insérer** > **Compléments**. Pour tous les types de complément, l’icône est également utilisée dans [AppSource](https://appsource.microsoft.com), si vous publiez votre complément dans AppSource.

L’image doit être dans un des formats de fichier suivants : GIF, JPG, PNG, EXIF, BMP ou TIFF. Pour les applications de volet de tâches et de contenu, l’image spécifiée doit contenir 32 x 32 pixels. Pour les applications de messagerie, la résolution d’image recommandée est de 64 x 64 pixels. Vous devez également spécifier une icône pour une utilisation avec les applications hôte Office en cours d’exécution sur des écrans haute résolution (DPI) à l’aide de l’élément [HighResolutionIconUrl](highresolutioniconurl.md). Pour plus d’informations, reportez-vous à la section _Créer une identité visuelle cohérente pour votre application_ dans [Création de listings efficaces dans AppSource et dans Office](/office/dev/store/create-effective-office-store-listings#create-a-consistent-visual-identity).

La modification de la valeur de l' `IconUrl` élément au moment de l’exécution n’est actuellement pas prise en charge.