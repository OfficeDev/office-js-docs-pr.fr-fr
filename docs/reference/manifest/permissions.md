---
title: Élément permissions dans le fichier manifest
description: L’élément permissions spécifie le niveau d’accès à l’API pour votre complément Office.
ms.date: 06/26/2020
localization_priority: Normal
ms.openlocfilehash: bc4cc2713d5a781c3407385470acd762910d17fd
ms.sourcegitcommit: 065bf4f8e0d26194cee9689f7126702b391340cc
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 07/01/2020
ms.locfileid: "45006457"
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

Pour plus d’informations, consultez la rubrique [demande d’autorisations pour l’utilisation d’API dans les compléments de contenu et du volet Office](../../develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md) et [Présentation des autorisations de complément Outlook](../../outlook/understanding-outlook-add-in-permissions.md).
