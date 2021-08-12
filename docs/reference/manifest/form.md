---
title: Élément Form dans le fichier manifeste
description: Paramètres UX pour les formulaires que votre complément de messagerie utilisera lors de l’exécution sur un appareil particulier (ordinateur de bureau, tablette ou téléphone).
ms.date: 04/09/2020
localization_priority: Normal
ms.openlocfilehash: f2a7265752cfcdd1030e4bcef36381692aeae1e8e1bb9d20f393c495f6beb48f
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/07/2021
ms.locfileid: "57086710"
---
# <a name="form-element"></a>Élément Form

Paramètres UX pour les formulaires que votre complément de messagerie utilisera lors de l’exécution sur un appareil particulier (ordinateur de bureau, tablette ou téléphone).

> [!IMPORTANT]
> Les éléments et les éléments sont disponibles uniquement dans les Outlook sur le web classiques (généralement connectés à des versions plus anciennes du serveur Exchange local) et Outlook `DesktopSettings` `TabletSettings` `PhoneSettings` 2013 sur Windows.

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

[FormSettings](formsettings.md)


## <a name="can-contain"></a>Peut contenir

|**Élément**|
|:-----|
|[DesktopSettings](desktopsettings.md)|
|[TabletSettings](tabletsettings.md)|
|[PhoneSettings](phonesettings.md)|
