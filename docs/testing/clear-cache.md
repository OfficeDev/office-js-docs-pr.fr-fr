---
title: Vider le cache Office
description: Découvrez comment effacer le cache Office sur votre ordinateur.
ms.date: 12/31/2019
localization_priority: Priority
ms.openlocfilehash: 3744d8125a5165569c262dc28622614853798c6f
ms.sourcegitcommit: d5ac9284d1e96dc91a9168d7641e44d88535e1a7
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 12/31/2019
ms.locfileid: "40915056"
---
# <a name="clear-the-office-cache"></a>Vider le cache Office

Vous pouvez supprimer un complément que vous avez précédemment chargé sur Windows, Mac ou iOS en vidant le cache Office sur votre ordinateur. 

En outre, si vous apportez des modifications au manifeste de votre complément (par exemple, vous mettez à jour le nom des fichiers d’icônes ou de texte de commandes du complément), videz le cache Office, puis rechargez le complément à l’aide d’un manifeste mis à jour. Cette action permettra à Office d’afficher le complément tel que décrit par le manifeste mis à jour.

## <a name="clear-the-office-cache-on-windows"></a>Vider le cache Office sur Windows

Pour vider le cache Office sur Windows, supprimez le contenu du dossier `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\`.

## <a name="clear-the-office-cache-on-mac"></a>Vider le cache Office sur Mac

[!include[additional cache folders on Mac](../includes/mac-cache-folders.md)]

##  <a name="clear-the-office-cache-on-ios"></a>Vider le cache Office sur iOS

Pour vider le cache Office sur iOS, appelez `window.location.reload(true)` à partir de JavaScript dans le complément pour forcer le rechargement. Vous pouvez également choisir de réinstaller Office.

## <a name="see-also"></a>Voir aussi

- [Manifeste XML des compléments Office](../develop/add-in-manifests.md)
- [Valider le manifeste d’un complément Office](troubleshoot-manifest.md)
- [Déboguer votre complément avec la journalisation runtime](runtime-logging.md)
- [Chargement de la version test des compléments Office](sideload-office-add-ins-for-testing.md)
- [Débogage des compléments Office](debug-add-ins-using-f12-developer-tools-on-windows-10.md)