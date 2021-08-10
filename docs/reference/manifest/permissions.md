---
title: Élément permissions dans le fichier manifest
description: L’élément Permissions spécifie le niveau d’accès d’API pour Office de votre application.
ms.date: 06/26/2020
localization_priority: Normal
ms.openlocfilehash: 2f2ccb4f6ec691b19cadea76a06520a9bad7a0b6c0e51699f2c8db67a3030de0
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/07/2021
ms.locfileid: "57089011"
---
# <a name="permissions-element"></a>Élément Permissions

Spécifie le niveau d’accès à l’API de votre complément Office ; vous devez demander des autorisations basées sur le principe des privilèges minimaux.

**Type de complément :** application de contenu, de volet Office, de messagerie (Mail)

## <a name="syntax"></a>Syntaxe

Pour les compléments du volet de tâches et de contenu :

```XML
 <Permissions> [Restricted | ReadDocument | ReadAllDocument | WriteDocument | ReadWriteDocument]</Permissions>
```

Pour les compléments de messagerie :

```XML
 <Permissions>[Restricted | ReadItem | ReadWriteItem | ReadWriteMailbox]</Permissions>
```

## <a name="contained-in"></a>Contenu dans

[OfficeApp](officeapp.md)

## <a name="remarks"></a>Remarques

Pour plus d’informations, voir Demande d’autorisations pour l’utilisation de l’API dans les [add-ins](../../develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md) de contenu et du volet Des tâches et Comprendre Outlook [autorisations de](../../outlook/understanding-outlook-add-in-permissions.md)votre application.
