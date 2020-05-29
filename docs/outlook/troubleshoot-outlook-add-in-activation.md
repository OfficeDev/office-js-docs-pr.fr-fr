---
title: Résolution des problèmes d’activation de complément contextuel Outlook
description: Si votre complément ne s’active pas comme prévu, vous devez rechercher dans les zones suivantes les raisons possibles.
ms.date: 05/27/2020
localization_priority: Normal
ms.openlocfilehash: 555ae2a45bf49d74d1fd439258fd87035644e86a
ms.sourcegitcommit: 77617f6ad06e07f5ff8078b26301748f73e2ee01
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/29/2020
ms.locfileid: "44413181"
---
# <a name="troubleshoot-outlook-add-in-activation"></a>Résolution des problèmes d’activation des compléments Outlook

L’activation de complément contextuel Outlook est basée sur les règles de l’activation dans le manifeste de complément. Lorsque les conditions de l’élément actuellement sélectionné respectent les règles d’activation du complément, l’application hôte active et affiche le bouton de complément dans l’interface utilisateur Outlook (volet de sélection de complément pour composer des compléments, barre de complément pour les compléments lus). Toutefois, si votre complément ne s’active pas comme prévu, vous devez rechercher les raisons possibles dans les zones suivantes.

## <a name="is-user-mailbox-on-a-version-of-exchange-server-that-is-at-least-exchange-2013"></a>Est-ce que la boîte aux lettres utilisateur se trouve sur une version d’Exchange Server correspondant au minimum à Exchange 2013 ?

En premier lieu, assurez-vous que le compte de messagerie utilisateur que vous employez pour le test se trouve sur une version d’Exchange Server correspondant au minimum à Exchange 2013. Si vous utilisez des fonctionnalités spécifiques ultérieures à Exchange 2013, assurez-vous que le compte utilisateur se trouve sur une version appropriée d’Exchange.

Vous pouvez vérifier la version d’Exchange 2013 en adoptant l’une des approches suivantes :

- Renseignez-vous auprès de votre administrateur Exchange Server.

- Si vous testez le complément sur Outlook sur le web ou sur appareils mobiles, dans un débogueur de script (par exemple le débogueur JScript disponible avec Internet Explorer), recherchez l’attribut **src** de la balise **script** qui spécifie l’emplacement à partir duquel les scripts sont chargés. Le chemin d’accès doit contenir une sous-chaîne **owa/15.0.516.x/owa2/...**, où **15.0.516.x** représente la version du serveur Exchange Server (par exemple **15.0.516.2**).

- Vous pouvez également utiliser la propriété [Office.context.mailbox.diagnostics.hostVersion](/javascript/api/outlook/office.diagnostics#hostversion) pour vérifier la version. Dans Outlook sur le web et sur appareils mobiles, cette propriété renvoie la version du serveur Exchange Server.

- Si vous pouvez tester le complément sur Outlook, servez-vous de cette technique de débogage simple, qui fait appel au modèle objet Outlook et à Visual Basic Editor :

    1. Tout d’abord, assurez-vous que les macros sont activées pour Outlook. Choisissez **Fichier**, **Options**, **Centre de gestion de la confidentialité**, **Paramètres du Centre de gestion de la confidentialité**, **Paramètres des macros**. Assurez-vous que l’option **Notifications pour toutes les macros** est sélectionnée dans le Centre de gestion de la confidentialité. Vous devez également avoir sélectionné **Activer les macros** au cours du démarrage d’Outlook.

    1. Sous l’onglet **Développeur** du ruban, choisissez **Visual Basic**.

       > [!NOTE]
       > Si vous ne voyez pas l’onglet **Développeur**, reportez-vous à la rubrique [Procédure : Afficher l’onglet Développeur sur le ruban](/visualstudio/vsto/how-to-show-the-developer-tab-on-the-ribbon) pour l’activer.

    1. Dans Visual Basic Editor, choisissez **Affichage**, **Fenêtre exécution**.

    1. Tapez ce qui suit dans la fenêtre Exécution pour afficher la version du serveur Exchange Server. La version principale de la valeur retournée doit être égale ou supérieure à 15.

       - S’il n’y a qu’un seul compte Exchange dans le profil de l’utilisateur :

       ```vb
        ?Session.ExchangeMailboxServerVersion
       ```

       - Si le même profil utilisateur comporte plusieurs comptes Exchange (`emailAddress` représente une chaîne qui contient l’adresse SMTP principale de l’utilisateur) :

       ```vb
        ?Session.Accounts.Item(emailAddress).ExchangeMailboxServerVersion
       ```

## <a name="is-the-add-in-disabled"></a>Le complément est-il désactivé ?

N’importe lequel des clients riches Outlook peut désactiver un complément pour des raisons de performances, notamment en cas de dépassement des seuils suivants : utilisation de l’UC ou de la mémoire, tolérance des incidents et durée nécessaire au traitement de toutes les expressions régulières pour un complément. Quand cela se produit, le client riche Outlook affiche une notification pour indiquer qu’il désactive le complément.

> [!NOTE]
> Seuls les clients riches Outlook surveillent l’utilisation des ressources. Toutefois, la désactivation d’un complément dans un client riche Outlook entraîne également la désactivation du complément dans Outlook sur le web et sur appareils mobiles.

Utilisez l’une des approches suivantes pour vérifier si un complément est désactivé :

- Dans Outlook sur le web, connectez-vous directement au compte de messagerie, choisissez l’icône Paramètres, puis choisissez **Gérer les compléments** afin d’accéder au Centre d’administration Exchange, où vous pouvez vérifier si le complément est activé.

- Dans Outlook sur Windows, accédez au mode Backstage, puis choisissez **Gérer les compléments**. Connectez-vous au Centre d’administration Exchange pour vérifier si le complément est activé.

- Dans Outlook sur Mac, choisissez **Gérer les compléments** dans la barre du complément. Connectez-vous au Centre d’administration Exchange pour vérifier si le complément est activé.

## <a name="does-the-tested-item-support-outlook-add-ins-is-the-selected-item-delivered-by-a-version-of-exchange-server-that-is-at-least-exchange-2013"></a>Les éléments testés prennent-ils en charge les compléments Outlook et sont-ils remis par une version d’Exchange Server correspondant au minimum à Exchange 2013 ?

Si votre complément Outlook est un complément de lecture et est supposé être activé lorsque l’utilisateur visualise un message (y compris les esmails, les demandes de réunions, les réponses et les annulations) ou un rendez-vous, même si ces éléments prennent en charge les compléments de manière générale, il existe des exceptions. Vérifiez si l’élément sélectionné est l’un de ceux[ répertoriés pour lesquels les compléments Outlook ne s’activent pas](outlook-add-ins-overview.md#mailbox-items-available-to-add-ins).

En outre, les rendez-vous étant toujours enregistrés au format RTF, une règle [ItemHasRegularExpressionMatch](../reference/manifest/rule.md#itemhasregularexpressionmatch-rule) qui spécifie une valeur **PropertyName** de **BodyAsHTML** n’active pas de complément pour un rendez-vous ou un message enregistré au format texte brut ou RTF.

Même si un élément de messagerie ne correspond pas à l’un des types ci-dessus, si cet élément n’a pas été remis par une version d’Exchange Server correspondant au minimum à Exchange 2013, les entités et les propriétés connues telles que l’adresse SMTP de l’expéditeur ne sont pas identifiées pour l’élément. Les règles d’activation qui dépendent de ces entités ou propriétés ne sont pas satisfaites et le complément n’est pas activé.

Si votre complément est un complément de composition et qu’il est censé être activé lorsque l’utilisateur compose un message ou une demande de réunion, assurez-vous que l’élément n’est pas protégé par IRM.

## <a name="is-the-add-in-manifest-installed-properly-and-does-outlook-have-a-cached-copy"></a>Est-ce que le manifeste du complément est correctement installé et est-ce qu’Outlook dispose d’une copie mise en cache ?

Ce scénario s’applique uniquement à Outlook sur Windows. Normalement, quand vous installez un complément Outlook pour une boîte aux lettres, le serveur Exchange copie le manifeste du complément de l’emplacement que vous indiquez vers la boîte aux lettres située sur ce serveur Exchange. Chaque fois qu’Outlook démarre, il lit l’ensemble des manifestes installés pour cette boîte aux lettres dans un cache temporaire situé à l’emplacement suivant :

```text
%LocalAppData%\Microsoft\Office\16.0\WEF
```

Par exemple, pour l’utilisateur John, le cache peut se trouver à C:\Users\john\AppData\Local\Microsoft\Office\16.0\WEF.

> [!IMPORTANT]
> Pour Outlook 2013 sur Windows, utilisez 15,0 au lieu de 16,0 pour l’emplacement :
>
> ```text
> %LocalAppData%\Microsoft\Office\15.0\WEF
> ```

Si un complément ne s’active pour aucun élément, cela peut signifier que le manifeste n’a pas été correctement installé sur le serveur Exchange ou qu’Outlook n’a pas lu correctement le manifeste au démarrage. À l’aide du Centre d’administration Exchange, assurez-vous que le complément est installé et activé pour votre boîte aux lettres, puis redémarrez le serveur Exchange, si nécessaire.

La figure 1 montre un résumé des étapes à suivre pour vérifier si Outlook dispose d’une version valide du manifeste.

**Figure 1 Organigramme des étapes à suivre pour vérifier si Outlook a correctement mis en cache le manifeste**

![Organigramme de vérification du manifeste](../images/troubleshoot-manifest-flow.png)

La procédure suivante décrit les détails.

1. Si vous modifiez le manifeste quand Outlook est ouvert et si vous n’utilisez pas Visual Studio 2012 ou une version ultérieure de Visual Studio pour développer le complément, désinstallez-le, puis réinstallez-le via le Centre d’administration Exchange.

1. Redémarrez Outlook, puis vérifiez si Outlook active désormais le complément.

1. Si Outlook n’active pas le complément, vérifiez si Outlook dispose d’une copie correctement mise en cache du manifeste du complément. Regardez dans le chemin d’accès suivant :

    ```text
    %LocalAppData%\Microsoft\Office\16.0\WEF
    ```

    Vous trouverez le manifeste dans le sous-dossier suivant :

    ```text
    \<insert your guid>\<insert base 64 hash>\Manifests\<ManifestID>_<ManifestVersion>
    ```

    > [!NOTE]
    > Voici un exemple d’un chemin d’accès à un manifeste installé pour une boîte aux lettres de l’utilisateur John :
    >
    > ```text
    > C:\Users\john\appdata\Local\Microsoft\Office\16.0\WEF\{8D8445A4-80E4-4D6B-B7AC-D4E6AF594E73}\GoRshCWa7vW8+jhKmyiDhA==\Manifests\b3d7d9d5-6f57-437d-9830-94e2aaccef16_1.2
    > ```

    Vérifiez si le manifeste du complément que vous testez figure parmi les manifestes mis en cache.

1. Si le manifeste est dans le cache, ignorez le reste de cette section, puis examinez les autres raisons possibles à la suite de cette section.

1. Si le manifeste n’est pas dans le cache, vérifiez si Outlook a réussi à lire le manifeste à partir du serveur Exchange Server. Pour ce faire, utilisez l’Observateur d’événements Windows :

    1. Sous **Journaux Windows**, choisissez **Application**.

    1. Recherchez un événement relativement récent pour lequel l’ID d’événement est égal à 63, ce qui correspond au téléchargement par Outlook d’un manifeste auprès d’un serveur Exchange Server.

    1. Si Outlook a réussi à lire un manifeste, l’événement journalisé doit présenter la description suivante :

        ```text
        The Exchange web service request GetAppManifests succeeded.
        ```

        Ignorez ensuite le reste de cette section, puis examinez les autres raisons possibles à la suite de cette section.

1. Si vous ne voyez pas d’événement réussi, fermez Outlook et supprimez tous les manifestes du chemin d’accès suivant :

    ```text
    %LocalAppData%\Microsoft\Office\16.0\WEF\<insert your guid>\<insert base 64 hash>\Manifests\
    ```

    Démarrez Outlook, puis vérifiez si Outlook active désormais le complément.

1. Si Outlook n’active pas le complément, revenez à l’étape 3 pour revérifier si Outlook a correctement lu le manifeste.

## <a name="is-the-add-in-manifest-valid"></a>Le manifeste du complément est-il valide ?

Consultez la rubrique relative à la [validation et à la résolution des problèmes de votre manifeste](../testing/troubleshoot-manifest.md) pour résoudre les problèmes de manifeste de complément.

## <a name="are-you-using-the-appropriate-activation-rules"></a>Utilisez-vous les règles d’activation appropriées ?

À partir de la version 1.1 du schéma des manifestes des Compléments Office, vous pouvez créer des compléments qui sont activés lorsque l’utilisateur se trouve dans un formulaire de composition (compléments de composition) ou de lecture (compléments de lecture). Assurez-vous que vous spécifiez les règles d’activation appropriées pour chaque type de formulaire dans lequel votre complément est censé être activé. Par exemple, vous ne pouvez activer des compléments de composition qu’à l’aide des règles [ItemIs](../reference/manifest/rule.md#itemis-rule) avec l’attribut **FormType** défini sur **Edit** ou **ReadOrEdit** et vous ne pouvez utiliser aucun autre type de règle, comme les règles [ItemHasKnownEntity](../reference/manifest/rule.md#itemhasknownentity-rule) et [ItemHasRegularExpressionMatch](../reference/manifest/rule.md#itemhasregularexpressionmatch-rule) pour les compléments de composition. Pour plus d’informations, voir [Règles d’activation pour les compléments Outlook](activation-rules.md).

## <a name="if-you-use-a-regular-expression-is-it-properly-specified"></a>Si vous utilisez une expression régulière, est-elle correctement spécifiée ?

Les expressions régulières contenues dans les règles d’activation font partie du fichier manifeste XML d’un complément de lecture. Si une expression régulière utilise certains caractères, veillez à bien suivre la séquence d’échappement correspondante prise en charge par les processeurs XML. Le tableau 1 répertorie ces caractères spéciaux.

**Tableau 1. Séquences d’échappement des expressions régulières**

|**Caractère**|**Description**|**Séquence d’échappement à utiliser**|
|:-----|:-----|:-----|
|`"`|Guillemets doubles|&amp;quot;|
|`&`|Esperluette|&amp;amp;|
|`'`|Apostrophe|&amp;apos;|
|`<`|Signe inférieur à|&amp;lt;|
|`>`|Signe supérieur à|&amp;gt;|

## <a name="if-you-use-a-regular-expression-is-the-read-add-in-activating-in-outlook-on-the-web-or-mobile-devices-but-not-in-any-of-the-outlook-rich-clients"></a>Si vous utilisez une expression régulière, est-ce que le complément de lecture s’active dans Outlook sur le web ou sur appareils mobiles, mais pas dans l’un des clients riches Outlook ?

Les clients riches Outlook utilisent un moteur d’expression régulière différent de celui utilisé par Outlook sur le web et sur appareils mobiles. Les clients riches Outlook utilisent le moteur d’expressions régulières C++ fourni avec la bibliothèque de modèles standard de Visual Studio. Ce moteur est conforme aux normes ECMAScript 5. Outlook sur le web et sur appareils mobiles utilisent l’évaluation d’expression régulière incluse dans JavaScript. Celle-ci est fournie par le navigateur et prend en charge un sur-ensemble d’ECMAScript 5.

Dans la plupart des cas, ces applications hôtes recherchent des correspondances identiques pour la même expression régulière dans une règle d’activation. Il existe cependant des exceptions. Par exemple, si l’expression régulière inclut une classe de caractères personnalisée basée sur des classes de caractères prédéfinies, un client riche Outlook peut renvoyer des résultats différents à partir d’Outlook sur le web et sur appareils mobiles. Par exemple, les classes de caractères qui contiennent des classes de caractères raccourcies `[\d\w]` renverraient des résultats différents. Dans ce cas, pour éviter d’obtenir des résultats différents sur différents hôtes, utilisez `(\d|\w)` à la place.

Testez minutieusement l’expression régulière. Si elle renvoie des résultats différents, réécrivez l’expression régulière pour qu’elle soit compatible avec les deux moteurs. Pour vérifier les résultats d’évaluation sur un client riche Outlook, écrivez un court programme C++ qui applique l’expression régulière par rapport à un échantillon du texte auquel vous essayez de la faire correspondre. Lors de son exécution dans Visual Studio, le programme de test C++ utilise la bibliothèque de modèles standards, simulant le comportement du client riche Outlook lors de l’exécution de la même expression régulière. Pour vérifier les résultats de l’évaluation sur Outlook sur le web ou sur appareils mobiles, utilisez le testeur d’expression régulière JavaScript privilégié.

## <a name="if-you-use-an-itemis-itemhasattachment-or-itemhasregularexpressionmatch-rule-have-you-verified-the-related-item-property"></a>Si vous utilisez une règle ItemIs, ItemHasAttachment ou ItemHasRegularExpressionMatch, avez-vous vérifié la propriété de l’élément connexe ?

Si vous utilisez une règle d’activation **ItemHasRegularExpressionMatch**, vérifiez si la valeur de l’attribut **PropertyName** correspond à ce que vous attendez pour l’élément sélectionné. Voici quelques conseils qui vous permettront de déboguer les propriétés correspondantes :

- Si l’élément sélectionné est un message et que vous spécifiez **BodyAsHTML** dans l’attribut **PropertyName**, ouvrez le message, puis choisissez **Afficher la source** afin de vérifier le corps du message dans la représentation HTML de cet élément.

- Si l’élément sélectionné est un rendez-vous ou si la règle d’activation spécifie **BodyAsPlaintext** dans l’élément **PropertyName**, vous pouvez utiliser le modèle objet Outlook et Visual Basic Editor dans Outlook sur Windows :

    1. Assurez-vous que les macros sont activées et que l’onglet **Développeur** s’affiche dans le ruban pour Outlook.

    1. Dans Visual Basic Editor, choisissez **Affichage**, **Fenêtre exécution**.

    1. Tapez ce qui suit pour afficher diverses propriétés en fonction du scénario.

        - Corps HTML de l’élément de message ou de rendez-vous sélectionné dans l’explorateur Outlook :

        ```vb
        ?ActiveExplorer.Selection.Item(1).HTMLBody
        ```
        - Corps en texte brut de l’élément de message ou de rendez-vous sélectionné dans l’explorateur Outlook :

        ```vb
        ?ActiveExplorer.Selection.Item(1).Body
        ```
        - Corps HTML de l’élément de message ou de rendez-vous ouvert dans l’inspecteur Outlook actif :

        ```vb
        ?ActiveInspector.CurrentItem.HTMLBody
        ```
        - Corps en texte brut de l’élément de message ou de rendez-vous ouvert dans l’inspecteur Outlook actif :

        ```vb
        ?ActiveInspector.CurrentItem.Body
        ```

Si la règle d’activation **ItemHasRegularExpressionMatch** spécifie **Subject** ou **SenderSMTPAddress**, ou si vous utilisez une règle **ItemIs** ou **ItemHasAttachment** et si vous êtes habitué à l’interface MAPI (ou si vous souhaitez l’utiliser), vous pouvez employer [MFCMAPI](https://github.com/stephenegriffin/mfcmapi) pour vérifier la valeur du tableau 2 dont dépend votre règle.

**Tableau 2. Règles d’activation et propriétés MAPI correspondantes**

|Type de règle|Vérifier cette propriété MAPI|
|:-----|:-----|
|Règle **ItemHasRegularExpressionMatch** avec **Subject**|[PidTagSubject](/office/client-developer/outlook/mapi/pidtagsubject-canonical-property)|
|Règle **ItemHasRegularExpressionMatch** avec **SenderSMTPAddress**|[PidTagSenderSmtpAddress](/office/client-developer/outlook/mapi/pidtagsendersmtpaddress-canonical-property) et [PidTagSentRepresentingSmtpAddress](/office/client-developer/outlook/mapi/pidtagsentrepresentingsmtpaddress-canonical-property)|
|**ItemIs**|[PidTagMessageClass](/office/client-developer/outlook/mapi/pidtagmessageclass-canonical-property)|
|**ItemHasAttachment**|[PidTagHasAttachments](/office/client-developer/outlook/mapi/pidtaghasattachments-canonical-property)|

Après avoir vérifié la valeur de propriété, vous pouvez utiliser un outil d’évaluation d’expression régulière pour vérifier si l’expression régulière trouve une correspondance dans cette valeur.

## <a name="does-the-host-application-apply-all-the-regular-expressions-to-the-portion-of-the-item-body-as-you-expect"></a>Est-ce que l’application hôte applique toutes les expressions régulières à la partie du corps de l’élément comme prévu ?

Cette section s’applique à toutes les règles d’activation qui utilisent des expressions régulières, notamment celles qui sont appliquées au corps de l’élément, dont la taille peut être importante et dont l’évaluation à la recherche de correspondances peut prendre plus de temps. Sachez que même si la valeur de la propriété de l’élément dont dépend une règle d’activation est celle attendue, l’application hôte ne pourra peut-être pas évaluer toutes les expressions régulières sur la valeur entière de la propriété de l’élément. Pour fournir des performances raisonnables et contrôler l’utilisation excessive des ressources par un complément de lecture, Outlook, Outlook sur le web et sur appareils mobiles répondent aux limites suivantes concernant le traitement des expressions régulières dans les règles d’activation au moment de l’exécution :

- Taille du corps d’élément évalué -- il existe des limites à la partie d’un corps d’élément pour lequel l’application hôte évalue une expression régulière. Ces limites dépendent de l’application hôte, du facteur de forme et du format du corps d’élément. Consultez les détails du tableau 2 dans [Limites d’activation et d’API JavaScript des compléments Outlook](limits-for-activation-and-javascript-api-for-outlook-add-ins.md).

- Nombre de correspondances d’expression régulière : les clients riches Outlook, Outlook sur le web et sur appareils mobiles renvoient chacun un nombre maximal de 50 correspondances d’expressions régulières. Ces correspondances sont uniques et les correspondances en double ne sont pas prises en compte par rapport à cette limite. Ne partez pas du principe que les correspondances renvoyées sont classées dans un ordre précis, ni que l’ordre dans un client riche Outlook est le même que celle dans Outlook sur le web et sur appareils mobiles. Si vous attendez de nombreuses correspondances pour des expressions régulières dans vos règles d’activation et qu’il manque une correspondance, il est possible que vous ayez dépassé cette limite.

- Longueur d’une correspondance d’expression régulière -- il existe des limites à la longueur d’une correspondance d’expression régulière retournée par l’application hôte. L’application hôte n’inclut aucune correspondance au-delà de la limite et n’affiche aucun message d’avertissement. Vous pouvez exécuter votre expression régulière à l’aide d’autres outils d’évaluation d’expression régulière ou via un programme de test autonome en C++ afin de vérifier s’il existe une correspondance qui dépasse les limites définies. Le tableau 3 récapitule ces limites. Pour plus d’informations, voir le tableau 3 dans [Limites d’activation et d’API JavaScript des compléments Outlook](limits-for-activation-and-javascript-api-for-outlook-add-ins.md).

    **Tableau 3. Limites de longueur pour une correspondance d’expression régulière**

    |Limite de longueur d’une correspondance d’expression régulière|Clients riches Outlook|Outlook sur le web ou sur appareils mobiles|
    |:-----|:-----|:-----|
    |Corps d’élément en texte brut|1,5 Ko|3 Ko|
    |Corps d’élément en HTML|3 Ko|3 Ko|

- Temps consacré à l’évaluation de toutes les expressions régulières d’un complément de lecture (pour un client riche Outlook) : par défaut, pour chaque complément de lecture, Outlook doit terminer l’évaluation de toutes les expressions régulières contenues dans ses règles d’activation en moins d’une seconde. Sinon, Outlook effectue jusqu’à trois nouvelles tentatives avant de désactiver le complément si l’évaluation ne peut pas être achevée. Outlook affiche un message dans la barre de notification pour indiquer que le complément a été désactivé. Vous pouvez modifier le délai disponible pour votre expression régulière en définissant une stratégie de groupe ou une clé de Registre. 

   > [!NOTE]
   > Si le client riche Outlook désactive un complément de lecture, celui-ci ne peut pas être utilisé pour la même boîte aux lettres sur le client riche Outlook, Outlook sur le web et sur appareils mobiles.

## <a name="see-also"></a>Voir aussi

- [Déployer et installer des compléments Outlook à des fins de test](testing-and-tips.md)
- [Règles d’activation pour les compléments Outlook](activation-rules.md)
- [Utiliser des règles d’activation d’expression régulière pour afficher un complément Outlook](use-regular-expressions-to-show-an-outlook-add-in.md)
- [Limites pour l’activation et l’API JavaScript pour les compléments Outlook](limits-for-activation-and-javascript-api-for-outlook-add-ins.md)
- [Valider et résoudre des problèmes avec votre manifeste](../testing/troubleshoot-manifest.md)
