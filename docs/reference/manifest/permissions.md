---
title: Élément permissions dans le fichier manifest
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 95cb45f89e2a5b92edc29bf32d0b47fcb2dbf8ce
ms.sourcegitcommit: a3ddfdb8a95477850148c4177e20e56a8673517c
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 02/20/2020
ms.locfileid: "42165544"
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

Pour plus d’informations, reportez-vous à la rubrique [demande d’autorisations pour l’utilisation des API dans les compléments](../../develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md) et [Présentation des autorisations de complément Outlook](../../outlook/understanding-outlook-add-in-permissions.md).
