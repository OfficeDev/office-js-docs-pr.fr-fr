---
title: Résolution des erreurs rencontrées par l’utilisateur avec des compléments Office
description: ''
ms.date: 06/20/2019
localization_priority: Priority
ms.openlocfilehash: 3222e8b7bc46608996c73284e2ee9b7c26c7afbe
ms.sourcegitcommit: 6d1cb188c76c09d320025abfcc99db1b16b7e37b
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/25/2019
ms.locfileid: "35226782"
---
# <a name="troubleshoot-user-errors-with-office-add-ins"></a>Résolution des erreurs rencontrées par l’utilisateur avec des compléments Office

Parfois, vos utilisateurs peuvent rencontrer des problèmes avec les compléments Office que vous développez. Par exemple, il se peut qu’un complément ne se charge pas ou soit inaccessible. Utilisez les informations de cet article pour résoudre les problèmes courants que vos utilisateurs rencontrent avec votre complément Office. 

Vous pouvez également utiliser [Fiddler](https://www.telerik.com/fiddler) pour identifier et déboguer les problèmes avec vos compléments.

## <a name="common-errors-and-troubleshooting-steps"></a>Erreurs courantes et étapes de dépannage

Le tableau suivant répertorie les messages d’erreur courants que les utilisateurs pourraient rencontrer, ainsi que les étapes que les utilisateurs peuvent suivre pour résoudre les erreurs.



|**Message d’erreur**|**Solution**|
|:-----|:-----|
|Erreur d’application : impossible d’accéder au catalogue|Vérifiez les paramètres de pare-feu. Le terme « catalogue » désigne AppSource. Ce message indique que l’utilisateur ne peut pas accéder à AppSource.|
|Erreur d’application : cette application n’a pas pu être démarrée. Fermez cette boîte de dialogue pour ignorer le problème, ou cliquez sur « Redémarrer » pour réessayer.|Vérifiez que les dernières mises à jour d’Office sont installés, ou téléchargez la [mise à jour pour Office 2013](https://support.microsoft.com/kb/2986156/).|
|Erreur : l’objet ne prend pas en charge la propriété ou la méthode « defineProperty »|Vérifiez qu’Internet Explorer ne fonctionne pas en mode de compatibilité. Accédez à Outils >  **Paramètres d’affichage de compatibilité**.|
|Désolé, nous n’avons pas pu charger l’application, car la version de votre navigateur n’est pas prise en charge. Cliquez ici pour obtenir la liste des versions de navigateur prises en charge.|Assurez-vous que le navigateur prend en charge le stockage local HTML5 ou réinitialisez les paramètres d’Internet Explorer. Pour plus d’informations sur les navigateurs pris en charge, reportez-vous à [Configuration requise pour exécuter des compléments Office](../concepts/requirements-for-running-office-add-ins.md).|


## <a name="outlook-add-in-doesnt-work-correctly"></a>§LTA Le complément Outlook ne fonctionne pas correctement

Si un complément Outlook s’exécutant sous Windows ne fonctionne pas correctement, essayez d’activer le débogage de script dans Internet Explorer. 


- Accédez à Outils >  **Options Internet** > **Avancées**.
    
- Sous  **Parcourir**, décochez les cases  **Désactiver le débogage des scripts (Internet Explorer)** et **Désactiver le débogage des scripts (autres applications)**.
    
Nous vous recommandons de décocher ces paramètres uniquement pour résoudre le problème. Si vous ne les réactivez pas, vous recevrez des invites. Une fois que le problème est résolu, recochez les cases  **Désactiver le débogage des scripts (Internet Explorer)** et **Désactiver le débogage des scripts (autres applications)**.


## <a name="add-in-doesnt-activate-in-office-2013"></a>Le complément ne s’active pas dans Office 2013

Le complément ne s’active pas lorsque l’utilisateur effectue les étapes suivantes :


1. connexion à son compte Microsoft dans Office 2013 ;
    
2. activation de la vérification à deux étapes pour son compte Microsoft ;
    
3. vérification de son identité après invitation lorsqu’il tente d’insérer un complément.
    
Pour résoudre ce problème, vérifiez que les dernières mises à jour Office sont installées ou téléchargez la [mise à jour pour Office 2013](https://support.microsoft.com/kb/2986156/).


## <a name="add-in-doesnt-load-in-task-pane-or-other-issues-with-the-add-in-manifest"></a>Le complément ne se charge pas dans le volet des tâches ou d’autres problèmes existent avec le manifeste du complément

Consultez la rubrique relative à la [validation et à la résolution des problèmes de votre manifeste](troubleshoot-manifest.md) pour déboguer le manifeste de votre complément.


## <a name="add-in-dialog-box-cannot-be-displayed"></a>La boîte de dialogue des compléments ne s’affiche pas

Lorsqu’un utilisateur utilise un complément Office, il est invité à autoriser l’affichage d’une boîte de dialogue. L’utilisateur choisit **Autoriser** et le message d’erreur suivant apparaît :

« Les paramètres de sécurité de votre navigateur nous empêchent de créer une boîte de dialogue. Essayez d’utiliser un autre navigateur, ou configurez votre navigateur de sorte que [URL] et le domaine affiché dans la barre d’adresse se trouvent dans la même zone de sécurité. »

![Capture d’écran du message d’erreur de la boîte de dialogue](http://i.imgur.com/3mqmlgE.png)

|**Navigateurs concernés**|**Plateformes concernées**|
|:--------------------|:---------------------|
|Internet Explorer, Microsoft Edge|Office sur le web|

Pour résoudre le problème, les utilisateurs finals et les administrateurs peuvent ajouter le domaine du complément à la liste des sites de confiance dans Internet Explorer. Utilisez la même procédure, que vous utilisiez le navigateur Internet Explorer ou Microsoft Edge.

> [!IMPORTANT]
> n’ajoutez pas l’URL d’un complément à votre liste de sites de confiance si vous ne faites pas confiance au complément.

Pour ajouter une URL à votre liste de sites de confiance :

1. Dans Internet Explorer, cliquez sur le bouton Outils et accédez à **Options Internet** > **Sécurité**.
2. Sélectionnez la zone **Sites de confiance**, puis choisissez **Sites**.
3. Entrez l’URL qui apparaît dans le message d’erreur, puis choisissez **Ajouter**.
4. Essayez d’utiliser le complément à nouveau. Si le problème persiste, vérifiez les paramètres pour les autres zones de sécurité et assurez-vous que le domaine du complément se trouve dans la même zone que l’URL qui s’affiche dans la barre d’adresse de l’application Office.

Ce problème se produit lorsque l’API de la boîte de dialogue est utilisée en mode contextuel. Pour éviter ce problème, utilisez l’indicateur [displayInFrame](/javascript/api/office/office.ui). Cela nécessite que votre page prenne en charge l’affichage dans un iframe. L’exemple suivant montre comment utiliser l’indicateur.

```js
Office.context.ui.displayDialogAsync(startAddress, {displayInFrame:true}, callback);
```

## <a name="changes-to-add-in-commands-including-ribbon-buttons-and-menu-items-do-not-take-effect"></a>Les modifications apportées aux commandes de complément, y compris les éléments de menu et les boutons du ruban ne s’appliquent pas

Si les modifications apportées au manifeste, par exemple aux noms de fichier des icônes de bouton dans le ruban ou au texte des éléments de menu, ne semblent pas appliquées, essayez de vider le cache Office de votre ordinateur. 

#### <a name="for-windows"></a>Pour Windows :
Supprimer le contenu du dossier `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\`.

#### <a name="for-mac"></a>Pour Mac :
Supprimer le contenu du dossier `~/Library/Containers/com.Microsoft.OsfWebHost/Data/`. 

[!include[additional cache folders on Mac](../includes/mac-cache-folders.md)]

#### <a name="for-ios"></a>Pour iOS :
Appelez `window.location.reload(true)` à partir de JavaScript dans le complément pour forcer le rechargement. Vous pouvez également choisir de réinstaller Office.

## <a name="see-also"></a>Voir aussi

- [Débogage de compléments dans Office sur le web](debug-add-ins-in-office-online.md) 
- [Charger une version test d’un complément Office sur iPad ou Mac](sideload-an-office-add-in-on-ipad-and-mac.md)  
- [Débogage des compléments Office sur iPad et Mac](debug-office-add-ins-on-ipad-and-mac.md)  
- [Valider et résoudre des problèmes avec votre manifeste](troubleshoot-manifest.md)
    
