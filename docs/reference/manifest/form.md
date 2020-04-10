---
title: Élément Form dans le fichier manifeste
description: Paramètres UX pour les formulaires que votre complément de messagerie utilisera lors de l’exécution sur un appareil particulier (ordinateur de bureau, tablette ou téléphone).
ms.date: 04/09/2020
localization_priority: Normal
ms.openlocfilehash: 3e8d60c13a72a50090075d7cd16a0719498c4982
ms.sourcegitcommit: c6e3bfd3deb77982d0b7082afd6a48678e96e1c3
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/10/2020
ms.locfileid: "43215067"
---
# <a name="form-element"></a>Élément Form

Paramètres UX pour les formulaires que votre complément de messagerie utilisera lors de l’exécution sur un appareil particulier (ordinateur de bureau, tablette ou téléphone).

> [!IMPORTANT]
> Les `DesktopSettings`éléments `TabletSettings`, et `PhoneSettings` sont disponibles uniquement dans les versions classiques d’Outlook sur le Web (généralement connectées à des versions plus anciennes de serveur Exchange local) et Outlook 2013 sur Windows.

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
