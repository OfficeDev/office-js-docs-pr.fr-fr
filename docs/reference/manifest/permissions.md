---
title: Élément permissions dans le fichier manifest
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: a70d72e454273873c6a30ffd82c3a2a5194f55e0
ms.sourcegitcommit: 350f5c6954dec3e9384e2030cd3265aaba7ae904
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 12/23/2019
ms.locfileid: "40851305"
---
# <a name="permissions-element"></a>Élément Permissions

Spécifie le niveau d’accès à l’API de votre complément Office ; vous devez demander des autorisations basées sur le principe des privilèges minimaux.

**Type de complément :** application de contenu, de volet Office, de messagerie (Mail)

## <a name="syntax"></a>Syntaxe

Pour les compléments du volet de tâches et de contenu:

```XML
 <Permissions> [Restricted | ReadDocument | ReadAllDocument | WriteDocument | ReadWriteDocument]</Permissions>
```

Pour les compléments de messagerie:

```XML
 <Permissions>[Restricted | ReadItem | ReadWriteItem | ReadWriteMailbox]</Permissions>
```

## <a name="contained-in"></a>Contenu dans

[OfficeApp](officeapp.md)

## <a name="remarks"></a>Remarques

Pour plus d’informations, reportez-vous à la rubrique [demande d’autorisations pour l’utilisation des API dans les compléments](/office/dev/add-ins/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins) et [Présentation des autorisations de complément Outlook](/outlook/add-ins/understanding-outlook-add-in-permissions).
