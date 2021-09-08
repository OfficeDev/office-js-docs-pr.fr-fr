---
title: Élément PhoneSettings dans le fichier manifeste
description: L’élément PhoneSettings spécifie l’emplacement source et les paramètres de contrôle qui s’appliquent lorsque votre module de messagerie est utilisé sur un téléphone.
ms.date: 04/09/2020
localization_priority: Normal
ms.openlocfilehash: d7957e23a77a0f837366e5cedc0e0f350b5635c8
ms.sourcegitcommit: 42c55a8d8e0447258393979a09f1ddb44c6be884
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/08/2021
ms.locfileid: "58938699"
---
# <a name="phonesettings-element"></a>PhoneSettings, élément

Spécifie l’emplacement source et les paramètres de contrôle qui s’appliquent lorsque votre complément de messagerie est utilisé sur un téléphone.

> [!IMPORTANT]
> L’élément est disponible uniquement dans les versions Outlook sur le web classiques (généralement connectées à des versions plus anciennes du serveur Exchange local) et Outlook `PhoneSettings` 2013 sur Windows. Pour prendre en charge Outlook sur Android et iOS, voir Les Outlook [Mobile](../../outlook/outlook-mobile-addins.md).

**Type de complément :** messagerie

## <a name="syntax"></a>Syntaxe

```XML
<Form xsi:type="ItemRead">
   <!--https://MyDomain.com/website.html is a placeholder for your own add-in website.-->
   <DesktopSettings>
      <!--If you opt to include RequestedHeight, it must be between 32px to 450px, inclusive.-->
      <RequestedHeight>360</RequestedHeight>
      <SourceLocation DefaultValue="https://MyDomain.com/website.html" />
   </DesktopSettings>
   <TabletSettings>
      <!--If you opt to include RequestedHeight, it must be between 32px to 450px, inclusive.-->
      <RequestedHeight>360</RequestedHeight>
      <SourceLocation DefaultValue="https://MyDomain.com/website.html" />
   </TabletSettings>
   <PhoneSettings>
      <SourceLocation DefaultValue="https://MyDomain.com/website.html" />
   </PhoneSettings>
</Form>
```

## <a name="contained-in"></a>Contenu dans

[Form](form.md)

