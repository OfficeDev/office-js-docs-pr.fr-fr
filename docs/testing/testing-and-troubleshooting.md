---
title: Résolution des erreurs rencontrées par l’utilisateur avec des compléments Office
description: Découvrez comment résoudre les erreurs des utilisateurs dans les compléments Office.
ms.date: 06/10/2022
ms.localizationpriority: medium
ms.openlocfilehash: 18bb3c180cd3af1eb8d045d7c69b9772532b04d4
ms.sourcegitcommit: 3abcf7046446e7b02679c79d9054843088312200
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 11/02/2022
ms.locfileid: "68810371"
---
# <a name="troubleshoot-user-errors-with-office-add-ins"></a>Résolution des erreurs rencontrées par l’utilisateur avec des compléments Office

At times your users might encounter issues with Office Add-ins that you develop. For example, an add-in fails to load or is inaccessible. Use the information in this article to help resolve common issues that your users encounter with your Office Add-in.

Vous pouvez également utiliser [Fiddler](https://www.telerik.com/fiddler) pour identifier et déboguer les problèmes avec vos compléments.

## <a name="common-errors-and-troubleshooting-steps"></a>Erreurs courantes et étapes de dépannage

Le tableau suivant répertorie les messages d’erreur courants que les utilisateurs pourraient rencontrer, ainsi que les étapes que les utilisateurs peuvent suivre pour résoudre les erreurs.

|**Message d’erreur**|**Solution**|
|:-----|:-----|
|Erreur d’application : impossible d’accéder au catalogue|Verify firewall settings."Catalog" refers to AppSource. This message indicates that the user cannot access AppSource.|
|APP ERROR: This app could not be started. Close this dialog to ignore the problem or click "Restart" to try again.|Vérifiez que les dernières mises à jour d’Office sont installés, ou téléchargez la [mise à jour pour Office 2013](https://support.microsoft.com/kb/2986156/).|
|Erreur : l’objet ne prend pas en charge la propriété ou la méthode « defineProperty »|Vérifiez qu’Internet Explorer ne fonctionne pas en mode de compatibilité. Accédez à **Outils** > **Paramètres de l’affichage de compatibilité**.|
|Désolé, nous n’avons pas pu charger l’application, car la version de votre navigateur n’est pas prise en charge. Cliquez ici pour obtenir la liste des versions de navigateur prises en charge.|Make sure that the browser supports HTML5 local storage, or reset your Internet Explorer settings. For information about supported browsers, see [Requirements for running Office Add-ins](../concepts/requirements-for-running-office-add-ins.md).|

## <a name="when-installing-an-add-in-you-see-error-loading-add-in-in-the-status-bar"></a>Lors de l’installation d’un complément, le message « Erreur lors du chargement du complément » s’affiche dans la barre d’état

1. Fermez Office.
1. Vérifiez que le manifeste est valide. Voir [Valider le manifeste d’un complément Office](troubleshoot-manifest.md).
1. Redémarrer le complément.
1. Réinstallez le complément.

Vous pouvez également nous adresser des commentaires : si vous utilisez Excel sur Windows ou Mac, vous pouvez adresser un commentaire à l’équipe chargée de l’extensibilité d’Office directement à partir d’Excel. Pour ce faire, sélectionnez **Fichier** > **Commentaires** > **Envoyer un smiley mécontent**. Envoyer un smiley mécontent fournit les journaux nécessaires pour comprendre le problème.

## <a name="outlook-add-in-doesnt-work-correctly"></a>Le complément Outlook ne fonctionne pas correctement

Si un complément Outlook s’exécutant sous Windows et [à l’aide d’Internet Explorer](../concepts/browsers-used-by-office-web-add-ins.md) ne fonctionne pas correctement, essayez d’activer le débogage de script dans Internet Explorer.

- Accédez à **Outils** > **Options** >  Internet **avancées**.
- Sous **Parcourir**, décochez les cases **Désactiver le débogage des scripts (Internet Explorer)** et **Désactiver le débogage des scripts (autres applications)**.

We recommend that you uncheck these settings only to troubleshoot the issue. If you leave them unchecked, you will get prompts when you browse. After the issue is resolved, check **Disable script debugging (Internet Explorer)** and **Disable script debugging (Other)** again.

## <a name="add-in-doesnt-activate-in-office-2013"></a>Le complément ne s’active pas dans Office 2013

Si le complément ne s’active pas lorsque l’utilisateur effectue les étapes suivantes.

1. connexion à son compte Microsoft dans Office 2013 ;

1. activation de la vérification à deux étapes pour son compte Microsoft ;

1. vérification de son identité après invitation lorsqu’il tente d’insérer un complément.

Pour résoudre ce problème, vérifiez que les dernières mises à jour Office sont installées ou téléchargez la [mise à jour pour Office 2013](https://support.microsoft.com/kb/2986156/).

## <a name="add-in-dialog-box-cannot-be-displayed"></a>La boîte de dialogue des compléments ne s’affiche pas

Lorsqu’un utilisateur utilise un complément Office, il est invité à autoriser l’affichage d’une boîte de dialogue. L’utilisateur choisit **Autoriser** et le message d’erreur suivant se produit.

"The security settings in your browser prevent us from creating a dialog box. Try a different browser, or configure your browser so that [URL] and the domain shown in your address bar are in the same security zone."

![Capture d’écran du message d’erreur de la boîte de dialogue.](../images/dialog-prevented.png)

|Navigateurs concernés|Plateformes concernées|
|:--------------------|:---------------------|
|Microsoft Edge|Office sur le web|

Pour résoudre le problème, les utilisateurs finaux ou les administrateurs peuvent ajouter le domaine du complément à la liste des sites approuvés dans le navigateur Microsoft Edge.

> [!IMPORTANT]
> n’ajoutez pas l’URL d’un complément à votre liste de sites de confiance si vous ne faites pas confiance au complément.

Pour ajouter une URL à votre liste de sites de confiance :

1. Dans **Panneau de configuration,** accédez à **Options Internet** > **Sécurité**.
1. Sélectionnez la zone **Sites de confiance**, puis choisissez **Sites**.
1. Entrez l’URL qui apparaît dans le message d’erreur, puis choisissez **Ajouter**.
1. Try to use the add-in again. If the problem persists, verify the settings for the other security zones and ensure that the add-in domain is in the same zone as the URL that is displayed in the address bar of the Office application.

This issue occurs when the Dialog API is used in pop-up mode. To prevent this issue from occurring, use the [displayInFrame](/javascript/api/office/office.ui) flag. This requires that your page support display within an iframe. The following example shows how to use the flag.

```js
Office.context.ui.displayDialogAsync(startAddress, {displayInIFrame:true}, callback);
```

## <a name="see-also"></a>Voir aussi

- [Résolution des erreurs rencontrées par l’utilisateur avec des compléments Office](troubleshoot-development-errors.md)
