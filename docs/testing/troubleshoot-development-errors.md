---
title: Résoudre les erreurs de développement avec les compléments Office
description: Découvrez comment résoudre les problèmes liés aux erreurs de développement dans les compléments Office.
ms.date: 09/08/2020
localization_priority: Normal
ms.openlocfilehash: 5801146165446352ec806f6f832e9976f96467ac
ms.sourcegitcommit: c6308cf245ac1bc66a876eaa0a7bb4a2492991ac
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/08/2020
ms.locfileid: "47409394"
---
# <a name="troubleshoot-development-errors-with-office-add-ins"></a>Résoudre les erreurs de développement avec les compléments Office

## <a name="add-in-doesnt-load-in-task-pane-or-other-issues-with-the-add-in-manifest"></a>Le complément ne se charge pas dans le volet des tâches ou d’autres problèmes existent avec le manifeste du complément

Voir [Valider le manifeste d’un complément Office](troubleshoot-manifest.md) et [Déboguer votre complément avec la journalisation runtime](runtime-logging.md) pour déboguer les problèmes de manifeste de compléments.

## <a name="changes-to-add-in-commands-including-ribbon-buttons-and-menu-items-do-not-take-effect"></a>Les modifications apportées aux commandes de complément, y compris les éléments de menu et les boutons du ruban ne s’appliquent pas

Si les modifications apportées au manifeste, par exemple aux noms de fichier des icônes de bouton dans le ruban ou au texte des éléments de menu, ne semblent pas appliquées, essayez de vider le cache Office de votre ordinateur. 

#### <a name="for-windows"></a>Pour Windows :

Supprimez le contenu du dossier `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\` et supprimez le contenu du dossier `%userprofile%\AppData\Local\Packages\Microsoft.Win32WebViewHost_cw5n1h2txyewy\AC\#!123\INetCache\` , s’il existe.

#### <a name="for-mac"></a>Pour Mac :

[!include[additional cache folders on Mac](../includes/mac-cache-folders.md)]

#### <a name="for-ios"></a>Pour iOS :
Appelez `window.location.reload(true)` à partir de JavaScript dans le complément pour forcer le rechargement. Vous pouvez également choisir de réinstaller Office.

## <a name="changes-to-static-files-such-as-javascript-html-and-css-do-not-take-effect"></a>Les modifications apportées aux fichiers statiques, tels que JavaScript, HTML et CSS ne sont pas prises en compte.

Le navigateur web met peut-être le contenu de ces fichiers en cache. Pour éviter cela, vous pouvez désactiver la mise en cache côté client lors du développement. Les spécifications dépendent du serveur utilisé. Dans la plupart des cas, elle implique l’ajout d’en-têtes aux réponses HTTP. Nous vous recommandons d’exécuter les actions suivantes :

- Cache-Control : « privé, aucun cache, aucun magasin »
- Pragma : « aucun cache »
- Date d’expiration : « -1 »

Un exemple d’opération dans un serveur Node.JS Express est disponible dans [ce fichier app.js](https://github.com/OfficeDev/Office-Add-in-NodeJS-SSO/blob/master/Complete/app.js). Un exemple de projet ASP.NET est disponible dans [ce fichier cshtml](https://github.com/OfficeDev/Office-Add-in-ASPNET-SSO/blob/master/Complete/Office-Add-in-ASPNET-SSO-WebAPI/Views/Shared/_Layout.cshtml).

Si votre complément est hébergé dans Internet Information Server (IIS), vous pouvez également ajouter ce qui suit à web. config.

```xml
<system.webServer>
  <staticContent>
    <clientCache cacheControlMode="UseMaxAge" cacheControlMaxAge="0.00:00:00" cacheControlCustom="must-revalidate" />
  </staticContent>
```

Si ces étapes ne semblent pas fonctionner au départ, vous devrez peut-être vider le cache du navigateur web. Effectuez cette opération à l’aide de l’interface utilisateur du navigateur web. Il est possible que le cache de périmètre ne soit pas correctement vidé lorsque vous essayez de le faire dans l’interface utilisateur Edge. Si cela se produit, exécutez la commande suivante dans l’invite de commandes Windows.

```bash
del /s /f /q %LOCALAPPDATA%\Packages\Microsoft.Win32WebViewHost_cw5n1h2txyewy\AC\#!123\INetCache\
```

## <a name="changes-made-to-property-values-dont-happen-and-there-is-no-error-message"></a>Les modifications apportées aux valeurs des propriétés ne se produisent pas et il n’y a pas de message d’erreur

Consultez la documentation de référence pour savoir si la propriété est en lecture seule. En outre, les définitions de la [machine à écrire](../develop/referencing-the-javascript-api-for-office-library-from-its-cdn.md) pour Office js spécifient les propriétés d’objet en lecture seule. Si vous tentez de définir une propriété en lecture seule, l’opération d’écriture échoue sans avertissement, sans qu’aucune erreur ne soit générée. L’exemple suivant tente à tort de définir la propriété en lecture seule [Chart.ID](/javascript/api/excel/excel.chart#id). Voir aussi [certaines propriétés ne peuvent pas être définies directement](../develop/application-specific-api-model.md#some-properties-cannot-be-set-directly).

```js
// This will do nothing, since `id` is a read-only property.
myChart.id = "5";
```

## <a name="add-in-doesnt-work-on-edge-but-it-works-on-other-browsers"></a>Le complément ne fonctionne pas sur Edge, mais fonctionne sur d’autres navigateurs

Consultez la rubrique [Troubleshooting Microsoft Edge Problems](../concepts/browsers-used-by-office-web-add-ins.md#troubleshooting-microsoft-edge-issues).

## <a name="excel-add-in-throws-errors-but-not-consistently"></a>Le complément Excel génère des erreurs, mais pas de façon cohérente

Consultez la rubrique [Troubleshoot Excel Add-ins](../excel/excel-add-ins-troubleshooting.md) pour obtenir les causes possibles.

## <a name="see-also"></a>Voir également

- [Débogage de compléments dans Office sur le web](debug-add-ins-in-office-online.md)
- [Charger une version test d’un complément Office sur iPad ou Mac](sideload-an-office-add-in-on-ipad-and-mac.md)  
- [Débogage des compléments Office sur iPad et Mac](debug-office-add-ins-on-ipad-and-mac.md)  
- [Complément Microsoft Office Extension de débogueur pour Visual Studio Code](debug-with-vs-extension.md)
- [Valider le manifeste d’un complément Office](troubleshoot-manifest.md)
- [Déboguer votre complément avec la journalisation runtime](runtime-logging.md)
- [Résolution des erreurs rencontrées par l’utilisateur avec des compléments Office](testing-and-troubleshooting.md)
