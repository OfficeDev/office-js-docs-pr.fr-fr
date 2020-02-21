---
title: Élément TabletSettings dans le fichier manifeste
description: ''
ms.date: 01/13/2020
localization_priority: Normal
ms.openlocfilehash: bf11dcfec4dfe2c40764722d23c7a69c289bba65
ms.sourcegitcommit: a3ddfdb8a95477850148c4177e20e56a8673517c
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 02/20/2020
ms.locfileid: "42163863"
---
# <a name="tabletsettings-element"></a>TabletSettings, élément

Spécifie les paramètres de contrôle qui s’appliquent lorsque votre complément de messagerie est utilisé sur une tablette.

> [!IMPORTANT]
> L' `TabletSettings` élément est disponible uniquement dans les versions classiques d’Outlook sur le Web (généralement connectées à des versions antérieures du serveur Exchange local) et Outlook 2013 sur Windows. Pour prendre en charge Outlook sur Android et iOS, reportez-vous à la rubrique [compléments pour Outlook Mobile](../../outlook/outlook-mobile-addins.md).

**Type de complément :** messagerie

## <a name="syntax"></a>Syntaxe

```XML
<Form xsi:type="ItemRead">
   <!--website.html is a placeholder for your own add-in website.-->
   <DesktopSettings>
      <SourceLocation DefaultValue="https://website.html" />
      <!--RequestedHeight must be between 240px to 800px, inclusive.-->
      <RequestedHeight>360</RequestedHeight>
   </DesktopSettings>
   <TabletSettings>
      <SourceLocation DefaultValue="https://website.html" />
      <!--RequestedHeight must be between 240px to 800px, inclusive.-->
      <RequestedHeight>360</RequestedHeight>
   </TabletSettings>
   <PhoneSettings>
      <SourceLocation DefaultValue="https://website.html" />
   </PhoneSettings>
</Form>
```

## <a name="contained-in"></a>Contenu dans

[Form](form.md)

