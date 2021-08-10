---
title: Résoudre les erreurs de développement avec Office de recherche
description: Découvrez comment résoudre les erreurs de développement dans les Office de développement.
ms.date: 06/11/2021
localization_priority: Normal
ms.openlocfilehash: a750f8db6e58406403d8bd0ef89e60128c2e08523375b4b2fbe6a904bfbae2d4
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/07/2021
ms.locfileid: "57093223"
---
# <a name="troubleshoot-development-errors-with-office-add-ins"></a>Résoudre les erreurs de développement avec Office de recherche

## <a name="add-in-doesnt-load-in-task-pane-or-other-issues-with-the-add-in-manifest"></a>Le complément ne se charge pas dans le volet des tâches ou d’autres problèmes existent avec le manifeste du complément

Voir [Valider le manifeste d’un complément Office](troubleshoot-manifest.md) et [Déboguer votre complément avec la journalisation runtime](runtime-logging.md) pour déboguer les problèmes de manifeste de compléments.

## <a name="changes-to-add-in-commands-including-ribbon-buttons-and-menu-items-do-not-take-effect"></a>Les modifications apportées aux commandes de complément, y compris les éléments de menu et les boutons du ruban ne s’appliquent pas

Si les modifications apportées au manifeste, par exemple aux noms de fichier des icônes de bouton dans le ruban ou au texte des éléments de menu, ne semblent pas appliquées, essayez de vider le cache Office de votre ordinateur. 

#### <a name="for-windows"></a>Pour Windows :

Supprimez le contenu du dossier `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\` et supprimez le contenu du `%userprofile%\AppData\Local\Packages\Microsoft.Win32WebViewHost_cw5n1h2txyewy\AC\#!123\INetCache\` dossier, s’il existe.

#### <a name="for-mac"></a>Pour Mac :

[!include[additional cache folders on Mac](../includes/mac-cache-folders.md)]

#### <a name="for-ios"></a>Pour iOS :

Appelez `window.location.reload(true)` à partir de JavaScript dans le complément pour forcer le rechargement. Vous pouvez également choisir de réinstaller Office.

## <a name="changes-to-static-files-such-as-javascript-html-and-css-do-not-take-effect"></a>Les modifications apportées aux fichiers statiques, tels que JavaScript, HTML et CSS ne sont pas prises en compte.

Le navigateur web met peut-être le contenu de ces fichiers en cache. Pour éviter cela, vous pouvez désactiver la mise en cache côté client lors du développement. Les spécifications dépendent du serveur utilisé. Dans la plupart des cas, elle implique l’ajout d’en-têtes aux réponses HTTP. Nous vous suggérons l’ensemble suivant.

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

## <a name="changes-made-to-property-values-dont-happen-and-there-is-no-error-message"></a>Les modifications apportées aux valeurs des propriétés ne se produisent pas et il n’y a aucun message d’erreur

Consultez la documentation de référence de la propriété pour voir si elle est en lecture seule. En outre, les [définitions TypeScript](../develop/referencing-the-javascript-api-for-office-library-from-its-cdn.md) pour Office JS spécifient quelles propriétés d’objet sont en lecture seule. Si vous tentez de définir une propriété en lecture seule, l’opération d’écriture échoue silencieusement, sans qu’aucune erreur ne soit lancée. L’exemple suivant tente par erreur de définir la propriété en lecture seule [Chart.id](/javascript/api/excel/excel.chart#id). Voir aussi [Certaines propriétés ne peuvent pas être définies directement.](../develop/application-specific-api-model.md#some-properties-cannot-be-set-directly)

```js
// This will do nothing, since `id` is a read-only property.
myChart.id = "5";
```

## <a name="getting-error-this-add-in-is-no-longer-available"></a>Obtention d’une erreur : « Ce module n’est plus disponible »

Voici quelques-unes des causes de cette erreur. Si vous découvrez d’autres causes, indiquez-nous l’outil de commentaires en bas de la page.

- Si vous utilisez Visual Studio, il se peut qu’il y a un problème avec le chargement de version secondaire. Fermez toutes les instances de l’hôte Office et des Visual Studio. Redémarrez Visual Studio puis réessayez d’appuyer sur F5.
- Le manifeste du add-in a été supprimé de son emplacement de déploiement, tel qu’un déploiement centralisé, un catalogue SharePoint ou un partage réseau.
- La valeur de [l’élément ID](../reference/manifest/id.md) dans le manifeste a été modifiée directement dans la copie déployée. Si, pour une raison quelconque, vous souhaitez modifier cet ID, supprimez d’abord le module de l’hôte Office, puis remplacez le manifeste d’origine par le manifeste modifié. Vous devez effacer le cache Office pour supprimer toutes les traces de l’original. Consultez la section Modifications apportées aux commandes [de add-in,](#changes-to-add-in-commands-including-ribbon-buttons-and-menu-items-do-not-take-effect) y compris les boutons du ruban et les éléments de menu ne prennent pas effet plus tôt dans cet article.
- Le manifeste du add-in a un qui n’est pas défini n’importe où dans la section Resources du manifeste, ou il y a une insmatance dans l’orthographe de l’endroit où il est utilisé et où il est défini dans la `resid` [](../reference/manifest/resources.md) `resid` `<Resources>` section.
- Il existe un `resid` attribut quelque part dans le manifeste avec plus de 32 caractères. Un attribut et l’attribut de la ressource correspondante dans la section ne peuvent pas être `resid` `id` plus de `<Resources>` 32 caractères.
- Le add-in possède une commande de add-in personnalisée, mais vous essayez de l’exécuter sur une plateforme qui ne les prend pas en charge. Pour plus d’informations, consultez les ensembles de conditions requises des commandes [de l’autre.](../reference/requirement-sets/add-in-commands-requirement-sets.md)

## <a name="add-in-doesnt-work-on-edge-but-it-works-on-other-browsers"></a>Le add-in ne fonctionne pas sur Edge, mais il fonctionne sur d’autres navigateurs

Voir [Résolution des problèmes Microsoft Edge problèmes.](../concepts/browsers-used-by-office-web-add-ins.md#troubleshooting-microsoft-edge-issues)

## <a name="excel-add-in-throws-errors-but-not-consistently"></a>Excel de l’équipe de sécurité envoie des erreurs, mais pas de manière cohérente

Pour [les causes possibles Excel résoudre les](../excel/excel-add-ins-troubleshooting.md) problèmes.

## <a name="manifest-schema-validation-errors-in-visual-studio-projects"></a>Erreurs de validation de schéma de manifeste dans Visual Studio projets

Si vous utilisez des fonctionnalités plus nouvelles qui nécessitent des modifications dans le fichier manifeste, vous pouvez obtenir des erreurs de validation dans Visual Studio. Par exemple, lorsque vous ajoutez l’élément pour implémenter le runtime JavaScript partagé, vous pouvez voir `<Runtimes>` l’erreur de validation suivante.

**L’élément « Host » dans l’espace de noms ' a un élément enfant non valide « Runtimes » dans http://schemas.microsoft.com/office/taskpaneappversionoverrides l’espace de noms http://schemas.microsoft.com/office/taskpaneappversionoverrides ' '**

Si cela se produit, vous pouvez mettre à jour les fichiers XSD Visual Studio aux dernières versions. Les versions de schéma les plus récentes sont à [l'[MS-APPENDIXMXML]: Appendix A: Full XML Schema](/openspecs/office_file_formats/ms-owemxml/c6a06390-34b8-4b42-82eb-b28be12494a8).

### <a name="locate-the-xsd-files"></a>Rechercher les fichiers XSD

1. Ouvrez votre projet dans Visual Studio.
1. Dans **l’Explorateur de** solutions, ouvrez manifest.xml fichier. Le manifeste se trouve généralement dans le premier projet sous votre solution.
1. Choose **View**  >  **Properties Window** (F4).
1. Dans la **fenêtre Propriétés,** choisissez les ellipses (...) pour ouvrir l’éditeur de **schémas XML.** Vous trouverez ici l’emplacement exact des dossiers de tous les fichiers de schéma que votre projet utilise.

### <a name="update-the-xsd-files"></a>Mettre à jour les fichiers XSD

1. Ouvrez le fichier XSD que vous souhaitez mettre à jour dans un éditeur de texte. Le nom de schéma de l’erreur de validation correspond au nom de fichier XSD. Par exemple, **ouvrez TaskPaneAppVersionOverridesV1_0.xsd**.
1. Recherchez le schéma mis à jour à [l’emplacement [MS-APPENDIXMXML]: Appendix A: Full XML Schema](/openspecs/office_file_formats/ms-owemxml/c6a06390-34b8-4b42-82eb-b28be12494a8). Par exemple, TaskPaneAppVersionOverridesV1_0 est au [schéma taskpaneappversionoverrides](/openspecs/office_file_formats/ms-owemxml/82e93ec5-de22-42a8-86e3-353c8336aa40).
1. Copiez le texte dans votre éditeur de texte.
1. Enregistrez le fichier XSD mis à jour.
1. Redémarrez Visual Studio pour récupérer les modifications apportées au nouveau fichier XSD.

Vous pouvez répéter le processus précédent pour tous les schémas supplémentaires qui ne sont pas à jour.

## <a name="see-also"></a>Voir également

- [Débogage de compléments dans Office sur le web](debug-add-ins-in-office-online.md)
- [Charger une version test d’un complément Office sur iPad ou Mac](sideload-an-office-add-in-on-ipad-and-mac.md)  
- [Débogage des compléments Office sur iPad et Mac](debug-office-add-ins-on-ipad-and-mac.md)  
- [Complément Microsoft Office Extension de débogueur pour Visual Studio Code](debug-with-vs-extension.md)
- [Valider le manifeste d’un complément Office](troubleshoot-manifest.md)
- [Déboguer votre complément avec la journalisation runtime](runtime-logging.md)
- [Résolution des erreurs rencontrées par l’utilisateur avec des compléments Office](testing-and-troubleshooting.md)
