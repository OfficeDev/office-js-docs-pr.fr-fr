---
title: Confidentialité et sécurité pour les compléments Office
description: Découvrez les aspects de confidentialité et de sécurité de la plateforme des compléments Office.
ms.date: 10/06/2020
localization_priority: Normal
ms.openlocfilehash: e7b301c6bbc28cdcb00d178d52b1c916a5e645a9
ms.sourcegitcommit: ceb8dd66f3fb9c963fce8446c2f6c65ead56fbc1
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 11/18/2020
ms.locfileid: "49132171"
---
# <a name="privacy-and-security-for-office-add-ins"></a>Confidentialité et sécurité pour les compléments Office

## <a name="understanding-the-add-in-runtime"></a>Présentation du runtime de complément

Les Compléments Office sont sécurisées par un environnement d’exécution de complément, un modèle d’autorisations à plusieurs niveaux et des gouverneurs de performances. Cette infrastructure protège l’expérience de l’utilisateur de la manière suivante.

- L’accès au cadre de l’interface utilisateur de l’application cliente Office est géré.

- Seul l’accès indirect au thread de l’interface utilisateur de l’application cliente Office est autorisé.

- Les interactions modales ne sont pas autorisées, par exemple, les appels vers JavaScript `alert` , `confirm` et les `prompt` fonctions ne sont pas autorisés, car ils sont modaux.

En outre, l’infrastructure d’exécution offre les avantages suivants pour s’assurer qu’un complément Office ne peut pas endommager l’environnement de l’utilisateur.

- Isole le processus dans lequel s’exécute le complément.

- Ne nécessite pas de remplacements de .dll ou de .exe, ni de composants ActiveX.

- Simplifie l’installation et la désinstallation des compléments.

De plus, l’utilisation des ressources de mémoire, de processeur et réseau par les compléments Office peut être régie afin de garantir de bonnes performances et une excellente fiabilité.

Les sections suivantes décrivent brièvement comment l’architecture d’exécution prend en charge l’exécution des compléments dans les clients Office sur les appareils Windows, Mac OS X et dans les navigateurs Web.

### <a name="clients-on-windows-and-os-x-devices"></a>Clients sur appareils Windows et OS X

Dans les clients pris en charge pour les ordinateurs de bureau et les tablettes, comme Excel sur Windows, et Outlook pour Windows et Mac, les compléments Office sont pris en charge en intégrant un composant in-process, le runtime des compléments Office, qui gère le cycle de vie du complément et permet l’interopérabilité entre le complément et l’application cliente. La page web du complément elle-même est hébergée hors processus. Comme indiqué dans la figure 1, sur un ordinateur de bureau ou une tablette, [la page web du complément est hébergée dans un contrôle Internet Explorer ou Microsoft Edge](browsers-used-by-office-web-add-ins.md)qui, à son tour, est hébergé dans un processus d’exécution du complément qui fournit la sécurité et l’isolation des performances.

Sur le bureau Windows, le mode protégé d’Internet Explorer doit être activé pour la zone de site sensible. En règle générale, il est activé par défaut. S’il est désactivé, une [erreur se produit](https://support.microsoft.com/help/2761180/apps-for-office-don-t-start-if-you-disable-protected-mode-for-the-restricted-sites-zone-in-internet-explorer) lorsque vous essayez de lancer un complément.

*Figure 1. Environnement d’exécution des compléments Office dans les clients de bureau et de tablette Windows*

![Diagramme illustrant une infrastructure client riche](../images/dk2-agave-overview-02.png)

Comme illustré dans la figure suivante, sur un ordinateur Mac OS X, la page Web du complément est hébergée dans un processus hôte d’exécution WebKit en bac à sable qui offre un niveau de sécurité et de protection des performances similaire.

*Figure 2. Environnement d’exécution des compléments Office dans les clients Mac OS X*

![Diagramme montrant les applications pour l’environnement d’exécution Office sur Mac OS X](../images/dk2-agave-overview-mac-02.png)

Le runtime des compléments Office gère les communications entre processus, la traduction des appels et des événements d’API JavaScript dans des appels et événements natifs, ainsi que la prise en charge à distance de l’interface utilisateur pour permettre au complément d’être restitué à l’intérieur du document, dans un volet Office ou de façon adjacente à un message électronique, une demande de réunion ou un rendez-vous.

### <a name="web-clients"></a>Clients web

Dans les clients Web pris en charge, les compléments Office sont hébergés dans un **IFRAME** qui s’exécute à l’aide de l’attribut HTML5 **sandbox** . Les composants ActiveX ou la navigation dans la page principale du client web ne sont pas autorisés. La prise en charge des compléments Office est activée dans les clients web par l’intégration de l’API JavaScript pour Office. Comme pour les applications clientes de bureau, l’API JavaScript gère le cycle de vie du complément et l’interopérabilité entre le complément et le client web. Cette interopérabilité est implémentée à l’aide d’une infrastructure spéciale de communication par publication de messages sur plusieurs cadres. La bibliothèque JavaScript (Office.js) utilisée sur les clients de bureau est disponible pour l’interaction avec le client web. La figure suivante illustre l’infrastructure qui prend en charge les compléments Office en cours d’exécution dans le navigateur, ainsi que les composants pertinents (le client Web, **IFRAME**, le runtime des compléments Office et l’interface API JavaScript pour Office) qui sont requis pour les prendre en charge.

*Figure 3. Infrastructure prenant en charge les compléments Office dans les clients web Office*

![Diagramme illustrant l’infrastructure du client Web](../images/dk2-agave-overview-03.png)

## <a name="add-in-integrity-in-appsource"></a>Intégrité de complément dans AppSource

Vous pouvez rendre vos compléments Office accessibles au public en les publiant dans AppSource. AppSource applique les mesures suivantes pour maintenir l’intégrité des compléments.

- Le serveur hôte d’un complément Office doit toujours utiliser le protocole SSL (Secure Sockets Layer) pour communiquer.

- Pour proposer des compléments, un développeur doit fournir la preuve de son identité, un accord contractuel et une stratégie de confidentialité conforme.

- Prend en charge un système d’évaluation par les utilisateurs pour les compléments disponibles afin de promouvoir une communauté exerçant une auto surveillance.

## <a name="optional-connected-experiences"></a>Expériences connectées facultatives

Les utilisateurs finaux et les administrateurs informatiques peuvent désactiver [expériences connectées facultatives dans ](/deployoffice/privacy/optional-connected-experiences) les clients de bureau et mobiles Office. Pour les compléments Office, l’impact de la désactivation du paramètre d' **expériences connectées facultatif** est que les utilisateurs ne peuvent plus accéder aux compléments ou à l’Office Store par le biais de ces clients. Toutefois, certains compléments Microsoft considérés comme essentiels ou stratégiques, et les compléments déployés par l’administrateur informatique d’une organisation via un [déploiement centralisé](../publish/centralized-deployment.md) seront toujours disponibles. De plus, les compléments et le magasin restent disponibles dans Outlook sur le Web, quel que soit l’état du paramètre.

Pour en savoir plus sur le comportement propre à Outlook, consultez la rubrique [confidentialité, autorisations et sécurité pour les compléments Outlook](../outlook/privacy-and-security.md#optional-connected-experiences).

Notez que si un administrateur informatique désactive l' [utilisation des expériences connectées dans Office](/deployoffice/privacy/manage-privacy-controls#policy-setting-for-most-connected-experiences), il a le même effet sur les compléments que la désactivation des expériences en connexion facultatives.

## <a name="addressing-end-users-privacy-concerns"></a>Réponse aux inquiétudes des utilisateurs finaux concernant la confidentialité

Cette section décrit la protection offerte par la plateforme des compléments Office du point de vue du client (utilisateur final) et vous donne des recommandations concernant la satisfaction des attentes des utilisateurs et la façon de gérer leurs informations d’identification personnelle (PII) en toute sécurité.

### <a name="end-users-perspective"></a>Point de vue des utilisateurs finaux

Les compléments Office sont créés à l’aide de technologies web qui sont exécutées dans un contrôle de navigateur ou un composant **iframe**. C’est la raison pour laquelle l’utilisation de compléments est semblable à la navigation sur les sites web, que ce soit sur Internet ou sur l’intranet. Les compléments peuvent être externes à une organisation (si le complément est acquis à partir d’AppSource) ou internes (si le complément est acquis à partir d’un catalogue de compléments Exchange Server, d’un catalogue d’applications SharePoint ou d’un partage de fichiers sur le réseau d’une organisation). Les compléments ont un accès limité au réseau et la plupart d’entre eux peuvent effectuer des opérations de lecture ou d’écriture dans le document ou l’élément de messagerie actif. La plateforme du complément applique certaines contraintes avant qu’un utilisateur ou un administrateur installe ou démarre ce complément. Mais, comme pour tout modèle d’extensibilité, les utilisateurs doivent faire preuve de prudence avant de lancer un complément inconnu.

La plateforme de compléments répond aux problèmes de confidentialité des utilisateurs finaux de la manière suivante.

- §LTA Les données communiquées avec le serveur web qui héberge un complément du volet Office, Outlook ou de contenu, ainsi que les communications entre le complément et tout service web, doivent toujours être chiffrées à l’aide du protocole SSL (Secure Socket Layer).

- Avant qu’un utilisateur n’installe un complément à partir d’AppSource, il peut afficher la politique de confidentialité et les conditions requises du complément. En outre, les compléments Outlook qui interagissent avec les boîtes aux lettres des utilisateurs exposent les autorisations spécifiques nécessaires ; l’utilisateur peut lire les conditions d’utilisation, les autorisations requises et la politique de confidentialité avant d’installer un complément Outlook.

- Lorsqu’ils partagent un document, les utilisateurs partagent également les compléments insérés dans ces documents ou qui y sont associés. Si un utilisateur ouvre un document qui contient un complément que l’utilisateur n’a pas utilisé auparavant, l’application cliente Office invite l’utilisateur à accorder l’autorisation pour que le complément s’exécute dans le document. Dans un environnement d’organisation, l’application cliente Office invite également l’utilisateur si le document provient d’une source externe.

- Les utilisateurs peuvent autoriser ou refuser l’accès à AppSource. Pour les compléments de contenu et du volet de tâches, les utilisateurs gèrent l’accès aux compléments et catalogues approuvés à partir du centre de gestion de la **confidentialité** sur le client Office hôte (ouvert à partir des options de **fichiers**  >  **Options**  >  paramètres du **Trust Center**  >  **Centre** de gestion de la confidentialité  >  **-catalogues de compléments approuvés**). Pour les compléments Outlook, les applications peuvent gérer les compléments en sélectionnant le bouton **gérer les compléments** : dans Outlook sur Windows, choisissez **fichier**  >  **gérer les compléments**. Dans Outlook sur Mac, cliquez sur le bouton **gérer les compléments** dans la barre de complément. Dans Outlook sur le web, choisissez le menu **Paramètres** (icône d’engrenage) > **Gérer les compléments**. Les administrateurs peuvent également gérer cet accès [à l’aide d’une stratégie de groupe](/previous-versions/office/office-2013-resource-kit/jj219429(v=office.15)#using-group-policy-to-manage-how-users-can-install-and-use-apps-for-office).

- La conception de la plateforme de complément offre une sécurité et des performances aux utilisateurs finaux de la manière suivante.

  - Un complément Office s’exécute dans un contrôle de navigateur Web hébergé dans un environnement d’exécution de complément distinct de l’application cliente Office. Cette conception fournit à la fois une isolation de la sécurité et des performances à partir de l’application cliente.

  - L’exécution dans un contrôle de navigateur web permet au complément de faire quasiment tout ce qu’une page web ordinaire exécutée dans un navigateur peut faire mais, en même temps, oblige le complément à suivre la stratégie d’origine identique pour l’isolation du domaine et les zones de sécurité.

Les compléments Outlook fournissent des fonctionnalités supplémentaires de sécurité et de performance grâce à une analyse de l’utilisation des ressources propres aux compléments Outlook. Pour en savoir plus, voir [Confidentialité, autorisations et sécurité pour les compléments Outlook](../outlook/privacy-and-security.md).

### <a name="developer-guidelines-to-handle-pii"></a>Recommandations à l’intention des développeurs en matière de gestion des PII

Vous trouverez ci-dessous une liste de conseils spécifiques pour la protection des données personnelles pour vous-même en tant que développeur de compléments Office.

- L’objet [Settings](/javascript/api/office/office.settings) est conçu pour conserver les paramètres de complément et les données d’état entre les sessions pour un complément de contenu ou du volet Office, mais il ne stocke pas les mots de passe et autres informations d’identification personnelle confidentielles dans l’objet **Settings**. Les données contenues dans l’objet **Settings** ne sont pas visibles par les utilisateurs finaux, mais elles sont stockées en tant que partie du format de fichier du document, qui est facilement accessible. Vous devez limiter l’utilisation par votre complément des informations d’identification personnelle et stocker celles qu’il exige sur le serveur hébergeant votre complément en tant que ressource sécurisée par l’utilisateur.

- Certaines applications peuvent exposer les informations d’identification personnelle dans le cadre de leur utilisation. Faites en sorte de stocker les données de vos utilisateurs de manière sécurisée, notamment l’identité, la situation géographique, les heures d’accès et autres informations d’identification, pour éviter que d’autres utilisateurs du complément puissent y accéder.

- Si votre complément est disponible dans AppSource, l’utilisation obligatoire de HTTPS dans AppSource assure la protection des informations d’identification personnelle transmises entre votre serveur web et l’ordinateur client ou l’appareil. Toutefois, si vous devez retransmettre ces données à d’autres serveurs, veillez à observer le même niveau de protection.

- Si vous stockez les informations d’identification personnelle des utilisateurs, veillez à en informer les utilisateurs et à leur permettre de les inspecter et de les supprimer. Si vous envoyez votre complément à AppSource, vous pouvez indiquer les données que vous collectez et l’utilisation qui en est faite dans la déclaration de confidentialité.

## <a name="developers-permission-choices-and-security-practices"></a>Choix des développeurs relatifs aux autorisations et aux pratiques de sécurité

Suivez les recommandations générales suivantes pour prendre en charge le modèle de sécurité des compléments Office et faire une exploration en détail pour chaque type de complément.

### <a name="permissions-choices"></a>Choix des autorisations

La plateforme de complément fournit un modèle d’autorisations que votre complément utilise pour déclarer le niveau d’accès aux données d’un utilisateur qui sont requises pour ses fonctionnalités. Chaque niveau d’autorisation correspond au sous-ensemble de l’interface API JavaScript pour Office que votre complément est autorisé à utiliser pour ses fonctionnalités. Par exemple, l’autorisation **WriteDocument** pour les compléments de contenu et du volet Office autorise l’accès à la méthode [document. setSelectedDataAsync](/javascript/api/office/office.document) qui permet à un complément d’écrire dans le document de l’utilisateur, mais n’autorise pas l’accès aux méthodes de lecture des données à partir du document. Ce niveau d’autorisation est utile pour les compléments qui doivent uniquement écrire dans un document, comme par exemple un complément où l’utilisateur peut requérir des données à insérer dans son document.

Nous vous recommandons vivement de demander des autorisations sur la base du  _principe de privilège minimal_. Autrement dit, vous ne devez demander l’autorisation d’accès qu’au sous-ensemble minimal de l’API que votre complément requiert pour fonctionner correctement. Par exemple, si votre complément a seulement besoin de lire des données dans le document d’un utilisateur pour ses fonctionnalités, vous ne devez pas demander plus que l’autorisation **ReadDocument**. (Gardez toutefois à l’esprit qu’en cas de demande d’autorisations insuffisantes, la plateforme du complément bloquera l’utilisation de certaines API par votre complément et des erreurs seront générées lors de l’exécution.)

Spécifiez des autorisations dans le manifeste de votre complément, comme montré dans l’exemple de la section ci-dessous, pour permettre aux utilisateurs de connaître le niveau d’autorisation requis pour un complément avant de décider de l’installer ou de l’activer pour la première fois. De plus, les compléments Outlook qui demandent l’autorisation **ReadWriteMailbox** requièrent des privilèges d’administrateur explicites pour l’installation.

L’exemple suivant montre comment un complément du volet Office spécifie l’autorisation **ReadDocument** dans son manifeste. À des fins de clarté par rapport aux autorisations, les autres éléments du manifeste ne sont pas affichés.

```xml
<?xml version="1.0" encoding="utf-8"?>
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.0"
           xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" 
           xmlns:ver="http://schemas.microsoft.com/office/appforoffice/1.0"
           xsi:type="TaskPaneApp">

... <!-- To keep permissions as the focus, not displaying other elements. -->
  <Permissions>ReadDocument</Permissions>
...
</OfficeApp>
```

Pour plus d’informations sur les autorisations pour les compléments de contenu et le volet des tâches, reportez-vous à la rubrique [Demande d’autorisations d’utilisation de l’API dans des compléments](../develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md).

Pour plus d’informations sur les autorisations pour les compléments Outlook, consultez les rubriques suivantes.

- [Confidentialité, autorisations et sécurité pour les compléments Outlook](../outlook/privacy-and-security.md)

- [Présentation des autorisations de complément Outlook](../outlook/understanding-outlook-add-in-permissions.md)

### <a name="same-origin-policy"></a>Stratégie d’origine identique

Étant donné que les compléments Office sont des pages Web qui s’exécutent dans un contrôle de navigateur Web, ils doivent suivre la stratégie d’origine identique appliquée par le navigateur. Par défaut, une page Web dans un domaine ne peut pas effectuer d’appels de service Web [XMLHttpRequest](https://www.w3.org/TR/XMLHttpRequest/) vers un autre domaine que celui où elle est hébergée.

Pour surmonter cette limitation, vous pouvez utiliser JSON/P--fournir un proxy pour le service Web en incluant une balise **script** avec un attribut **src** pointant vers un script hébergé sur un autre domaine. Vous pouvez créer au moyen d’un programme les balises **script**, en créant dynamiquement l’URL vers laquelle pointer l’attribut **src**, et en passant les paramètres à l’URL via les paramètres de requêtes de l’URI. Les fournisseurs de services web créent et hébergent du code JavaScript sur des URL spécifiques et renvoient des scripts différents selon les paramètres de requête URI. Ces scripts s’exécutent ensuite là où ils sont insérés et fonctionnent comme prévu.

§LTA Ci-dessous figure un exemple de JSON/P dans l’exemple de complément Outlook. 

```js
// Dynamically create an HTML SCRIPT element that obtains the details for the specified video.
function loadVideoDetails(videoIndex) {
    // Dynamically create a new HTML SCRIPT element in the webpage.
    var script = document.createElement("script");
    // Specify the URL to retrieve the indicated video from a feed of a current list of videos,
    // as the value of the src attribute of the SCRIPT element. 
    script.setAttribute("src", "https://gdata.youtube.com/feeds/api/videos/" + 
        videos[videoIndex].Id + "?alt=json-in-script&amp;callback=videoDetailsLoaded");
    // Insert the SCRIPT element at the end of the HEAD section.
    document.getElementsByTagName('head')[0].appendChild(script);
}
```

Exchange et SharePoint sont des proxys côté client qui permettent un accès sur plusieurs domaines. En général, la stratégie d’origine identique sur un intranet n’est pas aussi stricte que sur Internet. Pour plus d’informations, voir [Stratégie d’origine identique Partie 1 : Interdiction de regarder](/archive/blogs/ieinternals/same-origin-policy-part-1-no-peeking) et [Résolution des limites de stratégie d’origine identique dans les compléments Office](../develop/addressing-same-origin-policy-limitations.md).

### <a name="tips-to-prevent-malicious-cross-site-scripting"></a>Conseils pour éviter les scripts intersites malveillants

Un utilisateur mal intentionné pourrait attaquer l’origine d’un complément en entrant un script malveillant via le document ou les champs dans le complément. Un développeur doit traiter l’entrée de l’utilisateur pour éviter l’exécution du JavaScript d’un utilisateur malveillant dans son domaine. Voici quelques bonnes pratiques à suivre pour gérer l’entrée de l’utilisateur à partir d’un document ou d’un message électronique, ou via les champs d’un complément.

- Au lieu d’utiliser la propriété DOM [innerHTML](https://developer.mozilla.org/docs/Web/API/Element/innerHTML), utilisez les propriétés [innerText](https://developer.mozilla.org/docs/Web/API/Node/innerText) et [textContent](https://developer.mozilla.org/docs/DOM/Node.textContent) chaque fois que cela est possible. Procédez comme suit pour la prise en charge entre navigateurs Internet Explorer et Firefox.

    ```js
     var text = x.innerText || x.textContent
    ```

    Pour plus d’informations sur les différences entre **InnerText** et **TextContent**, voir [node. textContent](https://developer.mozilla.org/docs/DOM/Node.textContent). Pour plus d’informations sur la compatibilité DOM entre les navigateurs les plus répandus, voir les instructions relatives à la [compatibilité DOM W3C - HTML](https://www.quirksmode.org/dom/w3c_html.html#t07).

- Si vous devez utiliser **InnerHtml**, assurez-vous que l’entrée de l’utilisateur ne contient pas de contenu malveillant avant de la transmettre à **InnerHtml**. Pour plus d’informations et pour obtenir un exemple d’utilisation de **InnerHtml** en toute sécurité, reportez-vous à la rubrique [InnerHtml](https://developer.mozilla.org/docs/Web/API/Element/innerHTML) , propriété.

- Si vous utilisez jQuery, utilisez la méthode [.text()](https://api.jquery.com/text/) au lieu de la méthode [.html()](https://api.jquery.com/html/).

- Utilisez la méthode [toStaticHTML](https://developer.mozilla.org/en-US/docs/Web/HTML/Reference) pour supprimer les éléments et attributs HTML dynamiques des entrées des utilisateurs avant de les transmettre à **innerHTML**.

- Utilisez la fonction [encodeURIComponent](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/encodeuricomponent) ou [encodeURI](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/encodeuri) pour encoder le texte qui représente une URL ayant pour origine ou contenant une entrée utilisateur.

- Consultez les informations relatives au [développement de compléments sécurisés](/previous-versions/windows/apps/hh849625(v=win.10)) pour connaître d’autres meilleures pratiques en matière de création de solutions web plus sécurisées.

### <a name="tips-to-prevent-clickjacking"></a>Conseils pour éviter les « détournements de clic »

Étant donné que les compléments Office sont affichés dans un IFRAME lorsqu’ils sont exécutés dans un navigateur avec des applications clientes Office, utilisez les conseils suivants pour réduire le risque de [détournements](https://en.wikipedia.org/wiki/Clickjacking) --technique utilisée par les pirates pour tromper les utilisateurs en révélant des informations confidentielles.

Tout d’abord, identifiez les actions sensibles que votre complément est en mesure d’effectuer, notamment celles qu’un utilisateur non autorisé pourrait utiliser à des fins malveillantes, comme effectuer une opération financière ou publier des données sensibles. Par exemple, votre complément peut permettre à l’utilisateur d’envoyer un paiement à un destinataire qu’il a lui-même défini.

Ensuite, concernant ces opérations sensibles, votre complément doit demander à l’utilisateur de confirmer l’action avant que celle-ci ne soit exécutée. Cette confirmation doit décrire en détail les conséquences de l’action qui va être exécutée. De même, le cas échéant, elle doit indiquer à l’utilisateur comment empêcher que l’action soit exécutée au moyen d’un bouton spécifique portant la mention « Ne pas autoriser » ou en ignorant la confirmation.

Enfin, pour être certain qu’aucun pirate informatique ne peut être en mesure de cacher ou masquer la confirmation, vous devez afficher cette dernière en dehors du contexte du complément (c’est-à-dire pas dans une boîte de dialogue HTML).

Voici quelques exemples de la façon dont vous pouvez obtenir la confirmation.

- Envoyer à l’utilisateur un courrier électronique contenant un lien de confirmation.

- Envoyer à l’utilisateur un message texte contenant un code de confirmation qu’il peut saisir dans le complément.

- Ouvrir une boîte de dialogue de confirmation dans une nouvelle fenêtre de navigateur dirigeant vers une page qui ne peut pas être intégrée dans un iFrame. C’est généralement le modèle qui est utilisé par les pages de connexion. Utilisez l’[API de boîte de dialogue](../develop/dialog-api-in-office-add-ins.md) pour créer une boîte de dialogue.

Assurez-vous également que l’adresse que vous utilisez pour contacter l’utilisateur n’a pas pu être fournie par un pirate potentiel. Par exemple, pour les confirmations de paiement, utilisez l’adresse figurant dans le compte de l’utilisateur autorisé.

### <a name="other-security-practices"></a>Autres pratiques de sécurité

Les développeurs doivent également prendre note des pratiques de sécurité suivantes.

- Les développeurs ne doivent pas utiliser les contrôles ActiveX dans les compléments Office car les contrôles ActiveX ne prennent pas en charge la nature multiplateforme de la plateforme du complément.

- Les compléments de contenu et du volet des tâches adoptent les mêmes paramètres SSL que les paramètres par défaut dans le navigateur, ce qui permet à la plupart des contenus d’être fournis uniquement par SSL. Les compléments Outlook nécessitent que le contenu soit fourni par SSL. Les développeurs doivent spécifier dans l’élément **SourceLocation** du manifeste de complément une URL qui utilise le protocole HTTPS pour identifier l’emplacement du fichier HTML du complément.

  Pour vous assurer que les compléments ne livrent pas du contenu à l’aide du protocole HTTP, lors du test des compléments, les développeurs doivent s’assurer que les paramètres suivants sont sélectionnés dans **Options Internet** du **panneau de configuration** et qu’aucun avertissement de sécurité n’apparaît dans leurs scénarios de test.

  - Assurez-vous que le paramètre de sécurité, **Afficher un contenu mixte**, pour la zone **Internet** est défini sur **Demander**. Pour ce faire, sélectionnez l’une des options suivantes dans les **Options Internet**: sous l’onglet **sécurité** , sélectionnez la zone **Internet** , sélectionnez **niveau personnalisé**, faites défiler pour **afficher le contenu mixte** et sélectionnez **invite** s’il n’est pas déjà sélectionné.

  - Assurez-vous que l’option **Avertir en cas de changement entre mode sécurisé et non sécurisé** est sélectionnée sur l’onglet **Avancé** de la boîte de dialogue **Options Internet**.

- Afin que les compléments n’utilisent pas trop les ressources du processeur ou de la mémoire et provoquent un refus de services sur un ordinateur client, la plateforme établit des limites d’utilisation des ressources. Lors du test, les développeurs doivent vérifier si le complément fonctionne dans les limites d’utilisation des ressources.

- Avant de publier un complément, les développeurs doivent s’assurer que toutes les informations personnelles identifiables exposées dans les fichiers de leur complément sont sécurisées.

- Les développeurs ne devraient pas intégrer les clés qu’ils utilisent pour accéder aux API ou aux services tiers (tels que Bing, Google ou Facebook) directement dans les pages HTML de leur complément. À la place, ils doivent créer un service web personnalisé ou stocker les clés sous une autre forme de stockage web sécurisé qu’ils peuvent appeler pour passer la valeur de clé de leur complément.

- Les développeurs doivent effectuer ce qui suit lors de la soumission d’un complément à AppSource.

  - Héberger le complément qu’ils soumettent sur un serveur web qui prend en charge SSL.
  - Produire une déclaration énonçant une stratégie de confidentialité conforme.
  - Être prêts à signer un accord contractuel lorsqu’ils soumettent le complément.

Outre les règles d’utilisation des ressources, les développeurs de compléments Outlook doivent également s’assurer que leurs compléments respectent les limites de spécification des règles d’activation et l’utilisation de l’interface API JavaScript. Pour plus d’informations, voir [Limites pour l’activation et l’API JavaScript pour les compléments Outlook](../outlook/limits-for-activation-and-javascript-api-for-outlook-add-ins.md).

## <a name="it-administrators-control"></a>Contrôle des administrateurs informatiques

Dans un environnement d’entreprise, les administrateurs informatiques ont l’autorité ultime pour accorder ou refuser l’accès à AppSource et aux catalogues privés.

La gestion et l’application des paramètres Office s’effectuent avec des paramètres de stratégie de groupe. Vous pouvez les configurer via l’[outil de déploiement d’Office](/deployoffice/overview-of-the-office-2016-deployment-tool), conjointement avec l’[outil de personnalisation Office](/deployoffice/overview-of-the-office-customization-tool-for-click-to-run).

| Nom du paramètre | Description |
|--------------|-------------|
| Autoriser les compléments et les catalogues web non sécurisés | Permet aux utilisateurs d’exécuter des compléments non sécurisés, par exemple les compléments qui disposent d’une page web ou à des emplacements de catalogue qui ne sont pas sécurisées par SSL (https://) et ne figurent pas dans les zones Internet des utilisateurs. |
| Bloquer les compléments web | Vous permet d’empêcher les utilisateurs d’utiliser des compléments web. |
| Bloquer Office Store |  Vous permet d’empêcher les utilisateurs d’utiliser ou d’insérer des compléments web inclus dans Office Store. |

> [!IMPORTANT]
> Si vos groupes de travail utilisent plusieurs versions d’Office, les paramètres de stratégie de groupe doivent être configurés pour chaque version. Veuillez consulter la rubrique [Utiliser les stratégies de groupe pour gérer la manière dont les utilisateurs peuvent installer et utiliser des applications pour Office](/previous-versions/office/office-2013-resource-kit/jj219429(v=office.15)#using-group-policy-to-manage-how-users-can-install-and-use-apps-for-office) dans l’article [Vue d’ensemble des applications pour Office 2013](/previous-versions/office/office-2013-resource-kit/jj219429(v%3doffice.15)) pour plus d’informations sur les paramètres de stratégie de groupe pour Office 2013.

## <a name="see-also"></a>Voir aussi

- [Demande d’autorisations d’utilisation de l’API dans des compléments](../develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)
- [Confidentialité, autorisations et sécurité pour les compléments Outlook](../outlook/privacy-and-security.md)
- [Présentation des autorisations de complément Outlook](../outlook/understanding-outlook-add-in-permissions.md)
- [Limites pour l’activation et l’API JavaScript pour les compléments Outlook](../outlook/limits-for-activation-and-javascript-api-for-outlook-add-ins.md)
- [Résolutions des limites de stratégie d’origine identique dans les compléments Office](../develop/addressing-same-origin-policy-limitations.md)
- [Stratégie d’origine identique](https://www.w3.org/Security/wiki/Same_Origin_Policy)
- [Stratégie d’origine identique Partie 1 : Interdiction de regarder](/archive/blogs/ieinternals/same-origin-policy-part-1-no-peeking)
- [Stratégie d’origine identique pour JavaScript](https://developer.mozilla.org/docs/Web/Security/Same-origin_policy)
- [Mode de protection d’Internet Explorer](https://support.microsoft.com/help/2761180/apps-for-office-don-t-start-if-you-disable-protected-mode-for-the-restricted-sites-zone-in-internet-explorer)
- [Contrôles de confidentialité pour Microsoft 365 Apps](/deployoffice/privacy/overview-privacy-controls)
