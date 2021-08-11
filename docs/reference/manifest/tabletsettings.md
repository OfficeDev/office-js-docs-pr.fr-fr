---
title: Élément TabletSettings dans le fichier manifeste
description: L’élément TabletSettings spécifie les paramètres de contrôle qui s’appliquent lorsque votre module de messagerie est utilisé sur une tablette.
ms.date: 04/09/2020
localization_priority: Normal
ms.openlocfilehash: 3360a446396da058b5ced0127050d807f33e7bda007b1da88414b351782d0127
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/07/2021
ms.locfileid: "57094643"
---
# <a name="tabletsettings-element"></a>TabletSettings, élément

Spécifie les paramètres de contrôle qui s’appliquent lorsque votre complément de messagerie est utilisé sur une tablette.

> [!IMPORTANT]
> L’élément est disponible uniquement dans les versions Outlook sur le web classiques (généralement connectées à des versions plus anciennes du serveur Exchange local) et Outlook `TabletSettings` 2013 sur Windows. Pour prendre en charge Outlook sur Android et iOS, voir Les Outlook [Mobile](../../outlook/outlook-mobile-addins.md).

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
