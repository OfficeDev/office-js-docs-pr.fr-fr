---
title: Élément permissions dans le fichier manifest
description: L’élément Permissions spécifie le niveau d’accès d’API pour Office de votre application.
ms.date: 06/26/2020
ms.localizationpriority: medium
ms.openlocfilehash: a472d7a6f375c3a171fdd529b993aaf2c6109ce9
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/12/2021
ms.locfileid: "59153616"
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
