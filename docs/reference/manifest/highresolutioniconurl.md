---
title: Élément HighResolutionIconUrl dans le fichier manifeste
description: ''
ms.date: 12/04/2018
localization_priority: Normal
ms.openlocfilehash: 41008be6b60d260bef78808af2b8dee1fbd0864a
ms.sourcegitcommit: 5d29801180f6939ec10efb778d2311be67d8b9f1
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 02/27/2020
ms.locfileid: "42325268"
---
# <a name="highresolutioniconurl-element"></a>HighResolutionIconUrl, élément

Spécifie l’URL de l’image qui est utilisée pour représenter votre complément Office dans l’UX d’insertion UX et l’Office Store sur les écrans à haute résolution (DPI).

**Type de complément :** application de contenu, de volet Office, de messagerie

## <a name="syntax"></a>Syntaxe

```XML
<HighResolutionIconUrl DefaultValue="string" />
```

## <a name="can-contain"></a>Peut contenir

[Override](override.md)

## <a name="attributes"></a>Attributs

|**Attribut**|**Type**|**Obligatoire**|**Description**|
|:-----|:-----|:-----|:-----|
|DefaultValue|chaîne (URL)|obligatoire|Spécifie la valeur par défaut de ce paramètre, exprimée pour les paramètres régionaux spécifiés dans l’élément [DefaultLocale](defaultlocale.md).|

## <a name="remarks"></a>Remarques

Pour un complément de messagerie **, l’icône** > est affichée dans l’interface utilisateur**gérer les compléments** . Pour un complément de contenu ou de volet Office, l’icône s’affiche dans l’interface utilisateur, sous **Insérer** > **Compléments**.

L’image doit être dans un des formats de fichier suivants : GIF, JPG, PNG, EXIF, BMP ou TIFF. Pour les applications de contenu et de volet des tâches, la résolution d’image recommandée est de 64 x 64 pixels. Pour les applications de messagerie, l’image doit faire 128 x 128 pixels. Pour plus d’informations, voir la section _Créer une identité visuelle cohérente pour votre application_ dans [Création de listings efficaces dans AppSource et dans Office](/office/dev/store/create-effective-office-store-listings#create-a-consistent-visual-identity).
