---
title: Confidentialit? et s?curit? pour les compl?ments Office
description: ''
ms.date: 01/23/2018
ms.openlocfilehash: 326c8095b6ced105cc21492dc290a443212b3d3f
ms.sourcegitcommit: c72c35e8389c47a795afbac1b2bcf98c8e216d82
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/23/2018
---
# <a name="privacy-and-security-for-office-add-ins"></a>Confidentialit? et s?curit? pour les compl?ments Office

## <a name="understanding-the-add-in-runtime"></a>Pr?sentation du runtime de compl?ment

Les Compl?ments Office sont s?curis?es par un environnement d?ex?cution de compl?ment, un mod?le d?autorisations ? plusieurs niveaux et des gouverneurs de performances. Cette infrastructure prot?ge l?exp?rience utilisateur de la fa?on suivante : 

- L?acc?s ? l?infrastructure de l?interface utilisateur de l?application h?te est g?r?.

- Seul un acc?s indirect au thread de l?interface utilisateur de l?application h?te est autoris?.

- Les interactions modales ne sont pas autoris?es. Par exemple, les appels aux fonctions JavaScript **alert**, **confirm** et **prompt** ne sont pas autoris?s, car ils sont modaux.

En outre, l?infrastructure d?ex?cution offre les avantages suivants pour garantir qu?un compl?ment Office ne peut pas endommager l?environnement de l?utilisateur :

- Isole le processus dans lequel s?ex?cute le compl?ment.

- Ne n?cessite pas de remplacements de .dll ou de .exe, ni de composants ActiveX.

- Simplifie l?installation et la d?sinstallation des compl?ments.

De plus, l?utilisation des ressources de m?moire, de processeur et r?seau par les compl?ments Office peut ?tre r?gie afin de garantir de bonnes performances et une excellente fiabilit?. 

Les sections suivantes d?crivent bri?vement comment l?architecture d?ex?cution prend en charge l?ex?cution de compl?ments dans les clients Office sur des appareils Windows ou Mac OS X, et dans les clients Office Online sur le web.

> **REMARQUE**  Pour en savoir plus sur l?utilisation de la protection des informations Windows et d?Intune avec des compl?ments Office, reportez-vous ? l?article relatif ? l?[utilisation de la protection des informations Windows et d?Intune pour prot?ger des donn?es d?entreprise dans les documents utilisant des compl?ments Office](https://docs.microsoft.com/en-us/microsoft-365-enterprise/office-add-ins-wip).

### <a name="clients-for-windows-and-os-x-devices"></a>Clients pour les appareils Windows et OS X

Dans les clients pris en charge pour les ordinateurs de bureau et les tablettes, comme Excel, Outlook et Outlook pour Mac, les compl?ment Office sont pris en charge en int?grant un composant in-process, le runtime des compl?ments Office, qui g?re le cycle de vie du compl?ment et permet l?interop?rabilit? entre le compl?ment et l?application cliente. La page web du compl?ment elle-m?me est h?berg?e hors processus. Comme indiqu? dans la figure 1, sur un ordinateur de bureau ou une tablette, la page web du compl?ment est h?berg?e dans un contr?le Internet Explorer qui, ? son tour, est h?berg? dans un processus d?ex?cution du compl?ment qui fournit la s?curit? et l?isolation des performances.

Sur le bureau Windows, Le mode prot?g? d?Internet Explorer doit ?tre activ? pour la zone de site sensible. En r?gle g?n?rale, il est activ? par d?faut. S?il est d?sactiv?, une [erreur se produit](https://support.microsoft.com/en-us/help/2761180/apps-for-office-don-t-start-if-you-disable-protected-mode-for-the-restricted-sites-zone-in-internet-explorer) lorsque vous essayez de lancer un compl?ment.

*Figure 1. Environnement d?ex?cution des compl?ments Office dans les clients de bureau et de tablette Windows*

![Infrastructure de client riche](../images/dk2-agave-overview-02.png)

Comme indiqu? dans la figure suivante, sur un ordinateur de bureau Mac OS X, la page web du compl?ment est h?berg?e dans un processus h?te d?ex?cution Webkit en mode bac ? sable (sandbox) qui fournit un niveau similaire de s?curit? et de protection des performances. 

*Figure 2. Environnement d?ex?cution des compl?ments Office dans les clients Mac OS X*

![Applications pour environnement d'ex?cution Office sur OS X Mac](../images/dk2-agave-overview-mac-02.png)

Le runtime des compl?ments Office g?re les communications entre processus, la traduction des appels et des ?v?nements d?API JavaScript dans des appels et ?v?nements natifs, ainsi que la prise en charge ? distance de l?interface utilisateur pour permettre au compl?ment d??tre restitu? ? l?int?rieur du document, dans un volet Office ou de fa?on adjacente ? un message ?lectronique, une demande de r?union ou un rendez-vous.

### <a name="web-clients"></a>Clients web

Dans les clients web pris en charge, tels que Excel Online et Outlook Web App, les compl?ments Office sont h?berg?s dans un composant **iframe** ex?cut? ? l?aide de l?attribut HTML5 **sandbox**. Les composants ActiveX ou la navigation dans la page principale du client web ne sont pas autoris?s. La prise en charge des compl?ments Office est activ?e dans les clients web par l?int?gration de l?API JavaScript pour Office. Comme pour les applications clientes de bureau, l?API JavaScript g?re le cycle de vie du compl?ment et l?interop?rabilit? entre le compl?ment et le client web. Cette interop?rabilit? est impl?ment?e ? l?aide d?une infrastructure sp?ciale de communication par publication de messages sur plusieurs cadres. La biblioth?que JavaScript (Office.js) utilis?e sur les clients de bureau est disponible pour l?interaction avec le client web. La figure suivante illustre l?infrastructure qui prend en charge les compl?ments Office dans Office Online (sur navigateur) et les composants impliqu?s (client web, **iframe**, ex?cution des compl?ments Office et API JavaScript pour Office) qui sont requis pour les prendre en charge.

*Figure 3. Infrastructure prenant en charge les compl?ments Office dans les clients web Office*

![Infrastructure de client web](../images/dk2-agave-overview-03.png)

## <a name="add-in-integrity-in-appsource"></a>Int?grit? de compl?ment dans AppSource

Vous pouvez rendre vos compl?ments Office accessibles au public en les publiant dans AppSource. AppSource applique les mesures suivantes pour conserver l?int?grit? des compl?ments :


- Le serveur h?te d?un compl?ment Office doit toujours utiliser le protocole SSL (Secure Sockets Layer) pour communiquer.

- Pour proposer des compl?ments, un d?veloppeur doit fournir la preuve de son identit?, un accord contractuel et une strat?gie de confidentialit? conforme.

- Assurez-vous que le code source des compl?ments est accessible en lecture seule.

- Un syst?me de r?vision par les utilisateurs est pris en charge pour les compl?ments disponibles afin de promouvoir une communaut? d?autor?glementation.

## <a name="addressing-end-users-privacy-concerns"></a>R?ponse aux inqui?tudes des utilisateurs finaux concernant la confidentialit?

Cette section d?crit la protection offerte par la plateforme des compl?ments Office du point de vue du client (utilisateur final) et vous donne des recommandations concernant la satisfaction des attentes des utilisateurs et la fa?on de g?rer leurs informations d?identification personnelle (PII) en toute s?curit?.

### <a name="end-users-perspective"></a>Point de vue des utilisateurs finaux

Les compl?ments Office sont cr??s ? l?aide de technologies web qui sont ex?cut?es dans un contr?le de navigateur ou un composant **iframe**. C?est la raison pour laquelle l?utilisation de compl?ments est semblable ? la navigation sur les sites web, que ce soit sur Internet ou sur l?intranet. Les compl?ments peuvent ?tre externes ? une organisation (si le compl?ment est acquis ? partir d?AppSource) ou internes (si le compl?ment est acquis ? partir d?un catalogue de compl?ments Exchange Server, d?un catalogue de compl?ments SharePoint ou d?un partage de fichiers sur le r?seau d?une organisation). Les compl?ments ont un acc?s limit? au r?seau et la plupart d?entre eux peuvent effectuer des op?rations de lecture ou d??criture dans le document ou l??l?ment de messagerie actif. La plateforme du compl?ment applique certaines contraintes avant qu?un utilisateur ou un administrateur installe ou d?marre ce compl?ment. Mais, comme pour tout mod?le d?extensibilit?, les utilisateurs doivent faire preuve de prudence avant de lancer un compl?ment inconnu.

La plateforme du compl?ment r?pond aux inqui?tudes des utilisateurs finaux concernant la confidentialit? des mani?res suivantes :

- ?LTA Les donn?es communiqu?es avec le serveur web qui h?berge un compl?ment du volet Office, Outlook ou de contenu, ainsi que les communications entre le compl?ment et tout service web, doivent toujours ?tre chiffr?es ? l?aide du protocole SSL (Secure Socket Layer).

- Avant qu?un utilisateur n?installe un compl?ment ? partir d?AppSource, il peut afficher la politique de confidentialit? et les conditions requises du compl?ment. En outre, les compl?ments Outlook qui interagissent avec les bo?tes aux lettres des utilisateurs exposent les autorisations sp?cifiques n?cessaires ; l?utilisateur peut lire les conditions d?utilisation, les autorisations requises et la politique de confidentialit? avant d?installer un compl?ment Outlook.

- Lorsqu?ils partagent un document, les utilisateurs partagent ?galement les compl?ments ins?r?s dans ces documents ou qui y sont associ?s. Si un utilisateur ouvre un document qui contient un compl?ment qu?il n?a jamais utilis? auparavant, l?application h?te demande ? l?utilisateur d?accorder l?autorisation d?ex?cution du compl?ment dans le document. Dans un environnement d?entreprise, l?application h?te Office demande ?galement ? l?utilisateur si le document provient d?une source externe.

- Les utilisateurs peuvent activer ou d?sactiver l?acc?s ? AppSource. Pour les compl?ments de contenu et du volet Office, les utilisateurs g?rent l?acc?s aux compl?ments et aux catalogues approuv?s ? partir du **Centre de gestion de la confidentialit?** sur le client Office h?te (ouvert ? partir de **Fichier** > **Options** > **Centre de gestion de la confidentialit?** > **Param?tres du Centre de gestion de la confidentialit?** > **Catalogues de compl?ments approuv?s**). Pour les compl?ments Outlook, les utilisateurs peuvent g?rer les compl?ments en s?lectionnant le bouton **G?rer les compl?ments** ; dans Outlook pour Windows, choisissez **Fichier** > **G?rer les compl?ments**. Dans Outlook pour Mac, s?lectionnez le bouton **G?rer les compl?ments** dans la barre des compl?ments. Dans Outlook Web App, choisissez le menu **Param?tres**(ic?ne d?engrenage) > **G?rer les compl?ments**. Les administrateurs peuvent ?galement g?rer cet acc?s [? l?aide de la strat?gie de groupe](http://technet.microsoft.com/en-us/library/jj219429.aspx#BKMK_Managing).

- La conception de la plateforme du compl?ment offre s?curit? et performance aux utilisateurs finals des fa?ons suivantes :

  - Un compl?ment Office s?ex?cute dans un contr?le de navigateur web, qui est h?berg? dans un environnement d?ex?cution de compl?ments s?par? de l?application h?te Office. Cette conception offre ? la fois une s?curit? et une s?paration des performances de l?application h?te.

  - L?ex?cution dans un contr?le de navigateur web permet au compl?ment de faire quasiment tout ce qu?une page web ordinaire ex?cut?e dans un navigateur peut faire mais, en m?me temps, oblige le compl?ment ? suivre la strat?gie d?origine identique pour l?isolation du domaine et les zones de s?curit?.

Les compl?ments Outlook fournissent des fonctionnalit?s suppl?mentaires de s?curit? et de performance gr?ce ? une analyse de l?utilisation des ressources propres aux compl?ments Outlook. Pour en savoir plus, voir [Confidentialit?, autorisations et s?curit? pour les compl?ments Outlook](https://docs.microsoft.com/en-us/outlook/add-ins/privacy-and-security).

### <a name="developer-guidelines-to-handle-pii"></a>Recommandations ? l?intention des d?veloppeurs en mati?re de gestion des PII

Vous pouvez lire les recommandations g?n?rales en mati?re de protection de PII destin?es aux administrateurs informatiques et aux d?veloppeurs dans la rubrique [Protection de la confidentialit? des donn?es dans le d?veloppement et le test d?applications de gestion de ressources humaines](http://technet.microsoft.com/en-us/library/gg447064.aspx). Voici quelques recommandations en mati?re de protection de PII pour les d?veloppeurs de compl?ments Office :

- L?objet [Settings](https://dev.office.com/reference/add-ins/shared/settings) est con?u pour conserver les param?tres de compl?ment et les donn?es d??tat entre les sessions pour un compl?ment de contenu ou du volet Office, mais il ne stocke pas les mots de passe et autres informations d?identification personnelle confidentielles dans l?objet **Settings**. Les donn?es contenues dans l?objet **Settings** ne sont pas visibles par les utilisateurs finaux, mais elles sont stock?es en tant que partie du format de fichier du document, qui est facilement accessible. Vous devez limiter l?utilisation par votre compl?ment des informations d?identification personnelle et stocker celles qu?il exige sur le serveur h?bergeant votre compl?ment en tant que ressource s?curis?e par l?utilisateur.

- Certaines applications peuvent exposer les informations d?identification personnelle dans le cadre de leur utilisation. Faites en sorte de stocker les donn?es de vos utilisateurs de mani?re s?curis?e, notamment l?identit?, la situation g?ographique, les heures d?acc?s et autres informations d?identification, pour ?viter que d?autres utilisateurs du compl?ment puissent y acc?der.

- Si votre compl?ment est disponible dans AppSource, l?utilisation obligatoire de HTTPS dans AppSource assure la protection des informations d?identification personnelle transmises entre votre serveur web et l?ordinateur client ou l?appareil. Toutefois, si vous devez retransmettre ces donn?es ? d?autres serveurs, veillez ? observer le m?me niveau de protection.

- Si vous stockez les informations d?identification personnelle des utilisateurs, veillez ? en informer les utilisateurs et ? leur permettre de les inspecter et de les supprimer. Si vous envoyez votre compl?ment ? AppSource, vous pouvez indiquer les donn?es que vous collectez et l?utilisation qui en est faite dans la d?claration de confidentialit?.

## <a name="developers-permission-choices-and-security-practices"></a>Choix des d?veloppeurs relatifs aux autorisations et aux pratiques de s?curit?

Suivez les recommandations g?n?rales suivantes pour prendre en charge le mod?le de s?curit? des compl?ments Office et faire une exploration en d?tail pour chaque type de compl?ment.

### <a name="permissions-choices"></a>Choix des autorisations

La plateforme de compl?ment fournit un mod?le d?autorisations que votre compl?ment utilise pour d?clarer le niveau d?acc?s aux donn?es d?un utilisateur dont il a besoin pour ses fonctionnalit?s. Chaque niveau d?autorisation correspond au sous-ensemble de l?interface API JavaScript pour Office que votre compl?ment est autoris? ? utiliser dans le cadre de ses fonctionnalit?s. Par exemple, l?autorisation **WriteDocument** pour les compl?ments de contenu de volet Office donne l?acc?s ? la m?thode [Document.setSelectedDataAsync](https://dev.office.com/reference/add-ins/shared/document.setselecteddataasync) qui permet ? un compl?ment d??crire dans le document de l?utilisateur, mais ne permet pas l?acc?s ? l?une des m?thodes de lecture des donn?es du document. Ce niveau d?autorisation est utile pour les compl?ments qui ont seulement besoin d??crire dans un document, comme un compl?ment dans lequel l?utilisateur peut rechercher des donn?es ? ins?rer dans son document.

Nous vous recommandons vivement de demander des autorisations sur la base du  _principe de privil?ge minimal_. Autrement dit, vous ne devez demander l?autorisation d?acc?s qu?au sous-ensemble minimal de l?API que votre compl?ment requiert pour fonctionner correctement. Par exemple, si votre compl?ment a seulement besoin de lire des donn?es dans le document d?un utilisateur pour ses fonctionnalit?s, vous ne devez pas demander plus que l?autorisation **ReadDocument**. (Gardez toutefois ? l?esprit qu?en cas de demande d?autorisations insuffisantes, la plateforme du compl?ment bloquera l?utilisation de certaines API par votre compl?ment et des erreurs seront g?n?r?es lors de l?ex?cution.)

Sp?cifiez des autorisations dans le manifeste de votre compl?ment, comme montr? dans l?exemple de la section ci-dessous, pour permettre aux utilisateurs de conna?tre le niveau d?autorisation requis pour un compl?ment avant de d?cider de l?installer ou de l?activer pour la premi?re fois. De plus, les compl?ments Outlook qui demandent l?autorisation **ReadWriteMailbox** exigent des privil?ges d?administrateur explicites pour l?installation.

L?exemple suivant montre comment un compl?ment du volet Office sp?cifie l?autorisation  **ReadDocument** dans son manifeste. ? des fins de clart? par rapport aux autorisations, les autres ?l?ments du manifeste ne sont pas affich?s.

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

Pour plus d?informations sur les autorisations pour les compl?ments de contenu et du volet Office, reportez-vous ? la rubrique [Demande d?autorisations d?utilisation de l?API dans des compl?ments de contenu et de volet des t?ches](https://docs.microsoft.com/en-us/office/dev/add-ins/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins).

Pour plus d?informations sur les autorisations pour les compl?ments Outlook, reportez-vous aux rubriques suivantes :

- [Confidentialit?, autorisations et s?curit? pour les compl?ments Outlook](https://docs.microsoft.com/outlook/add-ins/privacy-and-security)

- [Pr?sentation des autorisations de compl?ment Outlook](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)

### <a name="same-origin-policy"></a>Strat?gie d?origine identique

Comme les compl?ments Office sont des pages web qui s?ex?cutent dans un contr?le de navigateur web, elles doivent suivre la strat?gie d?origine identique appliqu?e par le navigateur : par d?faut, une page web dans un domaine ne peut pas effectuer des appels de service web [XmlHttpRequest](http://www.w3.org/TR/XMLHttpRequest/) vers un domaine autre que celui o? il est h?berg?.

Pour contourner cette limitation, il est possible d?utiliser JSON/P -- JSON/P fournit un proxy pour le service web en incluant une balise  **script** avec un attribut **src** qui pointe vers un script h?berg? sur un autre domaine. Vous pouvez cr?er au moyen d?un programme les balises **script**, en cr?ant dynamiquement l?URL vers laquelle pointer l?attribut **src**, et en passant les param?tres ? l?URL via les param?tres de requ?tes de l?URI. Les fournisseurs de services web cr?ent et h?bergent du code JavaScript ? des URL sp?cifiques, et retournent diff?rents scripts en fonction des param?tres de requ?te de l?URI. Ces scripts s?ex?cutent ensuite ? l?emplacement o? ils ont ?t? ins?r?s et fonctionnent comme ils sont charg?s de le faire.

?LTA Ci-dessous figure un exemple de JSON/P dans l?exemple de compl?ment Outlook. 

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

Exchange et SharePoint sont des proxys c?t? client qui permettent un acc?s sur plusieurs domaines. En g?n?ral, la strat?gie d?origine identique sur un intranet n?est pas aussi stricte que sur Internet. Pour plus d?informations, voir [Strat?gie d?origine identique Partie 1 : Interdiction de regarder](http://blogs.msdn.com/b/ieinternals/archive/2009/08/28/explaining-same-origin-policy-part-1-deny-read.aspx) et [R?solution des limites de strat?gie d?origine identique dans les compl?ments Office](../develop/addressing-same-origin-policy-limitations.md).

### <a name="tips-to-prevent-malicious-cross-site-scripting"></a>Conseils pour ?viter les scripts intersites malveillants

Un utilisateur mal intentionn? pourrait attaquer l?origine d?un compl?ment en entrant un script malveillant dans le document ou des champs du compl?ment. Un d?veloppeur doit traiter les entr?es utilisateur pour ?viter d?ex?cuter le code JavaScript d'un utilisateur malveillant dans son domaine. Voici quelques bonnes pratiques ? suivre pour g?rer les entr?es utilisateur d?un document, d?un message ?lectronique ou via les champs d?un compl?ment :


- Au lieu d?utiliser la propri?t? DOM [innerHTML](http://msdn.microsoft.com/en-us/library/ie/ms533897.aspx), utilisez les propri?t?s [innerText](https://msdn.microsoft.com/library/ms533899.aspx) et [textContent](https://developer.mozilla.org/en-US/docs/DOM/Node.textContent) chaque fois que cela est possible. Utilisez ce qui suit afin d?assurer la prise en charge entre navigateurs pour Internet Explorer et Firefox :

    ```js
     var text = x.innerText || x.textContent
    ```

    Pour plus d?informations sur les diff?rences entre  **innerText** et **textContent**, voir [Node.textContent](https://developer.mozilla.org/en-US/docs/DOM/Node.textContent). Pour plus d?informations sur la compatibilit? DOM entre les navigateurs les plus r?pandus, voir les instructions relatives ? la [compatibilit? DOM W3C - HTML](http://www.quirksmode.org/dom/w3c_html.html#t07).

- Si vous devez utiliser  **innerHTML**, assurez-vous que l?entr?e de l?utilisateur ne comporte pas de contenu malveillant avant de le transmettre ?  **innerHTML**. Pour plus d?informations et pour obtenir un exemple montrant comment utiliser sans risque  **innerHTML**, voir la propri?t? [innerHTML](http://msdn.microsoft.com/en-us/library/ie/ms533897.aspx).

- Si vous utilisez jQuery, utilisez la m?thode [.text()](http://api.jquery.com/text/) au lieu de la m?thode [.html()](http://api.jquery.com/html/).

- Utilisez la m?thode [toStaticHTML](http://msdn.microsoft.com/en-us/library/ie/cc848922.aspx) pour supprimer les ?l?ments et attributs HTML dynamiques des entr?es des utilisateurs avant de les transmettre ? **innerHTML**.

- Utilisez la fonction [encodeURIComponent](http://msdn.microsoft.com/en-us/library/8202bce6-1342-40dc-a5ef-ac6d210a7d15.aspx) ou [encodeURI](http://msdn.microsoft.com/en-us/library/17bab5a2-bcd4-46c2-8b52-b2b5a0ed98a3.aspx) pour encoder le texte qui repr?sente une URL ayant pour origine ou contenant une entr?e utilisateur.

- Consultez les informations relatives au [d?veloppement de compl?ments s?curis?s](http://msdn.microsoft.com/en-us/library/windows/apps/hh849625.aspx) pour conna?tre d?autres meilleures pratiques en mati?re de cr?ation de solutions web plus s?curis?es.

### <a name="tips-to-prevent-clickjacking"></a>Conseils pour ?viter les ? d?tournements de clic ?

Comme les compl?ments Office sont restitu?s dans un IFrame lorsqu?ils sont ex?cut?s dans un navigateur avec les applications h?tes Office Online, suivez les conseils ci-dessous pour minimiser le risque de [d?tournement de clic](http://en.wikipedia.org/wiki/Clickjacking), une technique employ?e par les pirates informatiques pour inciter les internautes ? fournir des informations confidentielles.

Tout d?abord, identifiez les actions sensibles que votre compl?ment est en mesure d?effectuer, notamment celles qu?un utilisateur non autoris? pourrait utiliser ? des fins malveillantes, comme effectuer une op?ration financi?re ou publier des donn?es sensibles. Par exemple, votre compl?ment peut permettre ? l?utilisateur d?envoyer un paiement ? un destinataire qu?il a lui-m?me d?fini.

Ensuite, concernant ces op?rations sensibles, votre compl?ment doit demander ? l?utilisateur de confirmer l?action avant que celle-ci ne soit ex?cut?e. Cette confirmation doit d?crire en d?tail les cons?quences de l?action qui va ?tre ex?cut?e. De m?me, le cas ?ch?ant, elle doit indiquer ? l?utilisateur comment emp?cher que l?action soit ex?cut?e au moyen d?un bouton sp?cifique portant la mention ? Ne pas autoriser ? ou en ignorant la confirmation.

Enfin, pour ?tre certain qu?aucun pirate informatique ne peut ?tre en mesure de cacher ou masquer la confirmation, vous devez afficher cette derni?re en dehors du contexte du compl?ment (c?est-?-dire pas dans une bo?te de dialogue HTML).

Voici quelques exemples de m?thodes que vous pouvez utiliser pour obtenir la confirmation :

- Envoyer ? l?utilisateur un courrier ?lectronique contenant un lien de confirmation.

- Envoyer ? l?utilisateur un message texte contenant un code de confirmation qu?il peut saisir dans le compl?ment.

- Ouvrir une bo?te de dialogue de confirmation dans une nouvelle fen?tre de navigateur dirigeant vers une page qui ne peut pas ?tre int?gr?e dans un iFrame. C?est g?n?ralement le mod?le qui est utilis? par les pages de connexion. Utilisez l?[API de bo?te de dialogue](../develop/dialog-api-in-office-add-ins.md) pour cr?er une bo?te de dialogue.

Assurez-vous ?galement que l?adresse que vous utilisez pour contacter l?utilisateur n?a pas pu ?tre fournie par un pirate potentiel. Par exemple, pour les confirmations de paiement, utilisez l?adresse figurant dans le compte de l?utilisateur autoris?.

### <a name="other-security-practices"></a>Autres pratiques de s?curit?

Les d?veloppeurs doivent aussi tenir compte des pratiques de s?curit? suivantes :


- Les d?veloppeurs ne doivent pas utiliser les contr?les ActiveX dans les compl?ments Office car les contr?les ActiveX ne prennent pas en charge la nature multiplateforme de la plateforme du compl?ment.

- Les compl?ments de contenu et du volet Office adoptent les m?mes param?tres SSL que les param?tres par d?faut dans Internet Explorer, ce qui permet ? la plupart des contenus d??tre fournis uniquement par SSL. Les compl?ments Outlook n?cessitent que le contenu soit fourni par SSL. Les d?veloppeurs doivent sp?cifier dans l??l?ment **SourceLocation** du manifeste de compl?ment une URL qui utilise le protocole HTTPS pour identifier l?emplacement du fichier HTML du compl?ment.

    Pour s?assurer que les compl?ments ne d?livrent pas du contenu ? l?aide du protocole HTTP lors du test des compl?ments, les d?veloppeurs doivent s?assurer que les param?tres suivants sont s?lectionn?s dans Internet Explorer et qu?aucun avertissement de s?curit? n?appara?t dans leurs sc?narios de test :

    - V?rifiez que le param?tre de s?curit?, **Affiche un contenu mixte**, pour la zone **Internet** est d?fini sur **Demander**. Pour cela, proc?dez comme suit dans Internet Explorer : sur l?onglet **S?curit?** de la bo?te de dialogue **Options Internet**, s?lectionnez la zone **Internet**, s?lectionnez **Personnaliser le niveau**, recherchez **Afficher un contenu mixte**, et s?lectionnez **Demander** si l?option n?est pas d?j? s?lectionn?e.

    - Assurez-vous que l?option **Avertir en cas de changement entre mode s?curis? et non s?curis?** est s?lectionn?e sur l?onglet **Avanc?** de la bo?te de dialogue **Options Internet**.

- Afin que les compl?ments n?utilisent pas trop les ressources du processeur ou de la m?moire et provoquent un refus de services sur un ordinateur client, la plateforme ?tablit des limites d?utilisation des ressources. Lors du test, les d?veloppeurs doivent v?rifier si le compl?ment fonctionne dans les limites d?utilisation des ressources.

- Avant de publier un compl?ment, les d?veloppeurs doivent s?assurer que toutes les informations personnelles identifiables expos?es dans les fichiers de leur compl?ment sont s?curis?es.

- Les d?veloppeurs ne devraient pas int?grer les cl?s qu?ils utilisent pour acc?der aux API ou aux services tiers (tels que Bing, Google ou Facebook) directement dans les pages HTML de leur compl?ment. ? la place, ils doivent cr?er un service web personnalis? ou stocker les cl?s sous une autre forme de stockage web s?curis? qu?ils peuvent appeler pour passer la valeur de cl? de leur compl?ment.

- Les d?veloppeurs doivent proc?der comme suit lorsqu?ils envoient un compl?ment ? AppSource :

  - H?berger le compl?ment qu?ils soumettent sur un serveur web qui prend en charge SSL.
  - Produire une d?claration ?non?ant une strat?gie de confidentialit? conforme.
  - ?tre pr?ts ? signer un accord contractuel lorsqu?ils soumettent le compl?ment.

Outre les r?gles d?utilisation des ressources, les d?veloppeurs de compl?ments Outlook doivent ?galement s?assurer que leurs compl?ments respectent les limites de sp?cification des r?gles d?activation et l?utilisation de l?interface API JavaScript. Pour plus d?informations, voir [Limites pour l?activation et l?API JavaScript pour les compl?ments Outlook](http://msdn.microsoft.com/library/e0c9e3d0-517e-4333-b8bd-e169c51a07f6.aspx).

## <a name="it-administrators-control"></a>Contr?le des administrateurs informatiques

Dans un environnement d?entreprise, les administrateurs informatiques ont l?autorit? ultime pour accorder ou refuser l?acc?s ? AppSource et aux catalogues priv?s.

## <a name="see-also"></a>Voir aussi

- [Demande d?autorisations d?utilisation de l?API dans des compl?ments de contenu et de volet des t?ches](http://msdn.microsoft.com/library/da2efadc-4ebf-45fe-be39-397ac1eb1dbd.aspx)
- [Confidentialit?, autorisations et s?curit? pour les compl?ments Outlook](http://msdn.microsoft.com/library/44208fc4-05d4-42d8-ab20-faa89624de1c.aspx)
- [Pr?sentation des autorisations de compl?ment Outlook](https://docs.microsoft.com/en-us/outlook/add-ins/understanding-outlook-add-in-permissions)
- [Limites pour l?activation et l?API JavaScript pour les compl?ments Outlook](http://msdn.microsoft.com/library/e0c9e3d0-517e-4333-b8bd-e169c51a07f6.aspx)
- [R?solutions des limites de strat?gie d?origine identique dans les compl?ments Office](http://msdn.microsoft.com/library/36c800ae-1dda-4ea8-a558-37c89ffb161b.aspx)
- [Strat?gie d?origine identique](http://www.w3.org/Security/wiki/Same_Origin_Policy)
- [Strat?gie d?origine identique Partie 1 : Interdiction de regarder](http://blogs.msdn.com/b/ieinternals/archive/2009/08/28/explaining-same-origin-policy-part-1-deny-read.aspx)
- [Strat?gie d?origine identique pour JavaScript](https://developer.mozilla.org/En/Same_origin_policy_for_JavaScript)
- [Mode de protection d?Internet Explorer](https://support.microsoft.com/en-us/help/2761180/apps-for-office-don-t-start-if-you-disable-protected-mode-for-the-restricted-sites-zone-in-internet-explorer)
