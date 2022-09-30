---
title: Glossaire des termes des compléments Office
description: Glossaire des termes couramment utilisés dans la documentation des compléments Office.
ms.date: 09/14/2022
ms.localizationpriority: medium
ms.openlocfilehash: ef8df6e344698f7d67ebe7afe1759e13630b385d
ms.sourcegitcommit: cff5d3450f0c02814c1436f94cd1fc1537094051
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/30/2022
ms.locfileid: "68234913"
---
# <a name="office-add-ins-glossary"></a>Glossaire des compléments Office

Il s’agit d’un glossaire des termes couramment utilisés dans la documentation des compléments Office.

## <a name="add-in"></a>macro complémentaire

Les compléments Office sont des applications web qui étendent les applications Office. Ces applications web ajoutent de nouvelles fonctionnalités à l’application Office, telles que l’ajout de données externes, l’automatisation des processus ou l’incorporation d’objets interactifs dans des documents Office.

Les compléments Office diffèrent des compléments VBA, COM et VSTO, car ils offrent une prise en charge multiplateforme (généralement web, Windows, Mac et iPad) et sont basés sur des technologies web standard (HTML, CSS et JavaScript). Le langage de programmation principal d’un complément Office est JavaScript ou TypeScript.

## <a name="add-in-commands"></a>commandes de complément

**Les commandes de complément sont des** éléments d’interface utilisateur, tels que des boutons et des menus, qui étendent l’interface utilisateur Office pour votre complément. Lorsque les utilisateurs sélectionnent un élément de commande de complément, ils lancent des actions telles que l’exécution de code JavaScript ou l’affichage du complément dans un volet Office. Les commandes de complément permettent à votre complément de ressembler à une partie d’Office, ce qui donne aux utilisateurs plus de confiance dans votre complément. Pour en savoir plus [, consultez les commandes de complément pour Excel, PowerPoint et Word](../design/add-in-commands.md) et les [commandes de complément pour Outlook](../outlook/add-in-commands-for-outlook.md) .

Voir aussi : [ruban, bouton du ruban](#ribbon-ribbon-button).

## <a name="application"></a>application

**L’application** fait référence à une application Office. Les applications Office qui prennent en charge les compléments Office sont Excel, OneNote, Outlook, PowerPoint, Project et Word.

Voir aussi : [client](#client), [hôte](#host), [application Office, client Office](#office-application-office-client).

## <a name="application-specific-api"></a>API spécifique à l’application

Les API spécifiques à l’application fournissent des objets fortement typés qui interagissent avec des objets natifs d’une application Office spécifique. Par exemple, vous appelez les API JavaScript Excel pour accéder aux feuilles de calcul, plages, tableaux, graphiques, etc. Les API spécifiques à l’application sont actuellement disponibles pour Excel, OneNote, PowerPoint, Visio et Word. Pour en savoir plus, consultez le [modèle d’API spécifique à l’application](../develop/application-specific-api-model.md) .

Voir aussi : [API courante](#common-api).

## <a name="client"></a>Client

**Le client** fait généralement référence à une application Office. Les applications Office ou les clients qui prennent en charge les compléments Office sont Excel, OneNote, Outlook, PowerPoint, Project et Word.

Voir aussi : [application](#application), [hôte](#host), [application Office, client Office](#office-application-office-client).

## <a name="common-api"></a>API communes

Les API courantes sont utilisées pour accéder à des fonctionnalités telles que l’interface utilisateur, les dialogues et les paramètres clients qui sont courants dans plusieurs applications Office. Ce modèle d’API utilise des[rappels](https://developer.mozilla.org/docs/Glossary/Callback_function), qui vous permettent de spécifier une seule opération dans chaque demande envoyée à l’application Office.

Les API courantes ont été introduites avec Office 2013 et sont utilisées pour interagir avec Office 2013 ou version ultérieure. Certaines API courantes sont des API héritées du début des années 2010. Excel, PowerPoint et Word ont tous des fonctionnalités d’API communes, mais la plupart de ces fonctionnalités ont été remplacées ou remplacées par le modèle d’API propre à l’application. Les API spécifiques à l’application sont recommandées lorsque cela est possible.

D’autres API courantes, telles que les API communes liées à Outlook, à l’interface utilisateur et à l’authentification, sont les API modernes et préférées à ces fins. Pour plus d’informations sur le modèle objet d’API common, consultez [Common JavaScript API object model](../develop/office-javascript-api-object-model.md).

Voir aussi : [API spécifique à l’application](#application-specific-api).

## <a name="content-add-in"></a>complément de contenu

**Les compléments de contenu sont des** vues web ou des vues de navigateur web qui sont incorporées directement dans des documents Excel, OneNote ou PowerPoint. Les compléments de contenu permettent aux utilisateurs d’accéder aux contrôles d’interface qui exécutent le code pour modifier des documents ou afficher des données d’une source de données. Utilisez les compléments de contenu lorsque vous souhaitez incorporer des fonctionnalités directement dans le document. Pour en savoir plus, consultez les [compléments Office de contenu](../design/content-add-ins.md) .

Voir aussi : [webview](#webview).

## <a name="content-delivery-network-cdn"></a>réseau de distribution de contenu (CDN)

Un **réseau de distribution de contenu** ou **CDN** est un réseau distribué de serveurs et de centres de données. Il offre généralement une disponibilité et des performances des ressources plus élevées par rapport à un seul serveur ou centre de données.

## <a name="contoso"></a>Contoso

**Contoso** Ltd. (également appelée Contoso et Contoso University) est une société fictive utilisée par Microsoft comme exemple de société et de domaine.

## <a name="custom-function"></a>fonction personnalisée

Une **fonction personnalisée** est une fonction définie par l’utilisateur qui est empaqueté avec un complément Excel. Les fonctions personnalisées permettent aux développeurs d’ajouter de nouvelles fonctions, au-delà des fonctionnalités Excel classiques, en définissant ces fonctions en JavaScript dans le cadre d’un complément. Les utilisateurs d’Excel peuvent accéder aux fonctions personnalisées comme ils le feraient pour n’importe quelle fonction native dans Excel. Pour plus d’informations, consultez [Créer des fonctions personnalisées dans Excel](../excel/custom-functions-overview.md) .

## <a name="custom-functions-runtime"></a>runtime de fonctions personnalisées

Un **runtime de fonctions personnalisées** est un [runtime JavaScript uniquement](../testing/runtimes.md#javascript-only-runtime) qui exécute des fonctions personnalisées sur certaines combinaisons d’hôtes et de plateformes Office. Il n’a pas d’interface utilisateur et ne peut pas interagir avec Office.js API. Si votre complément a uniquement des fonctions personnalisées, il s’agit d’un bon runtime léger à utiliser. Si vos fonctions personnalisées doivent interagir avec le volet Office ou les API Office.js, configurez un [runtime partagé](../testing/runtimes.md#shared-runtime). Pour plus d’informations, consultez [Configurer votre complément Office pour utiliser un runtime partagé](../develop/configure-your-add-in-to-use-a-shared-runtime.md) .

Voir aussi : [runtime](#runtime), [runtime partagé](#shared-runtime).

## <a name="custom-functions-only-add-in"></a>complément de fonctions personnalisées uniquement

Complément qui contient une fonction personnalisée, mais aucune interface utilisateur telle qu’un volet Office. Les fonctions personnalisées de ce type de complément s’exécutent dans un [runtime JavaScript uniquement](../testing/runtimes.md#javascript-only-runtime). Une fonction personnalisée qui inclut une interface utilisateur peut utiliser un runtime partagé ou une combinaison d’un runtime JavaScript uniquement et d’un runtime html. Nous vous recommandons d’utiliser un runtime partagé si vous disposez d’une interface utilisateur.

Voir aussi : [fonction personnalisée](#custom-function), [runtime de fonctions personnalisées](#custom-functions-runtime).

## <a name="host"></a>host

**\<Host\>** fait généralement référence à une application Office. Les applications Office ou hôtes qui prennent en charge les compléments Office sont Excel, OneNote, Outlook, PowerPoint, Project et Word.

Voir aussi : [application](#application), [client](#client), [application Office, client Office](#office-application-office-client).

## <a name="office-application-office-client"></a>Application Office, client Office

**Le client Office** fait référence à une application Office. Les applications Office ou les clients qui prennent en charge les compléments Office sont Excel, OneNote, Outlook, PowerPoint, Project et Word.

Voir aussi : [application](#application), [client](#client), [hôte](#host).

## <a name="perpetual"></a>Perpétuel

**Perpetual** fait référence aux versions d’Office disponibles via un contrat de licence en volume ou des canaux de vente au détail.

D’autres contenus Microsoft peuvent utiliser le terme **non-abonnement** pour représenter ce concept.

Voir aussi : [vente au détail, vente au détail perpétuelle](#retail-retail-perpetual), [licence en volume, licence en volume perpétuelle, licence en volume](#volume-licensed-volume-licensed-perpetual-volume-licensing)

## <a name="platform"></a>platform

Une **plateforme** fait généralement référence au système d’exploitation exécutant l’application Office. Les plateformes qui prennent en charge les compléments Office incluent windows, Mac, iPad et navigateurs web.

## <a name="quick-start"></a>démarrage rapide

Un **démarrage rapide** est une description générale des compétences clés et des connaissances requises pour le fonctionnement de base d’un programme particulier. Dans la documentation des compléments Office, un démarrage rapide est une introduction au développement d’un complément pour une application particulière, telle qu’Outlook. Un démarrage rapide contient une série d’étapes qu’un développeur de compléments peut effectuer en environ 5 minutes, ce qui entraîne un complément fonctionnel et un environnement de développement fonctionnel.

Voir aussi : [didacticiel](#tutorial).

## <a name="requirement-set"></a>ensemble de conditions requises

[!include[Requirement set note](../includes/office-js-requirement-sets.md)]

## <a name="retail-retail-perpetual"></a>retail, retail perpetual

**La vente au détail** fait référence aux versions perpétuelles d’Office disponibles via les canaux de vente au détail. Celles-ci n’incluent pas les versions fournies par un abonnement Microsoft 365 ni un contrat de licence en volume.

D’autres contenus Microsoft peuvent utiliser le terme **achat** unique ou **consommateur** pour représenter ce concept.

Voir aussi : [perpétuel](#perpetual)

## <a name="ribbon-ribbon-button"></a>ruban, bouton du ruban

Un **ruban** est une barre de commandes qui organise les fonctionnalités d’une application en une série d’onglets ou de boutons en haut d’une fenêtre. Un **bouton de ruban** est l’un des boutons de cette série. Pour plus [d’informations, consultez Afficher ou masquer le ruban dans Office](https://support.microsoft.com/office/d946b26e-0c8c-402d-a0f7-c6efa296b527#ID0EBBD=Newer_Versions) .

## <a name="runtime"></a>Duree

Un **runtime** est l’environnement hôte (y compris un moteur JavaScript et généralement également un moteur de rendu HTML) dans lequel le complément s’exécute. Dans Office sur Windows et Office sur Mac, le runtime est un contrôle de navigateur incorporé (ou webview) tel qu’Internet Explorer, Edge Hérité, Edge WebView2 ou Safari. Différentes parties d’une exécution de complément dans des runtimes distincts. Par exemple, les commandes de complément, les fonctions personnalisées et le code du volet Office utilisent généralement des runtimes distincts, sauf si vous configurez un [runtime partagé](../testing/runtimes.md#shared-runtime). Pour plus d’informations, consultez [Runtimes in Office Add-ins](../testing/runtimes.md) and [Browsers used by Office Add-ins](../concepts/browsers-used-by-office-web-add-ins.md) .

Voir aussi : [runtime de fonctions personnalisées](#custom-functions-runtime), [runtime partagé](#shared-runtime), [webview](#webview).

## <a name="shared-runtime"></a>runtime partagé

Un **runtime partagé** permet à tout le code de votre complément, y compris le volet Office, les commandes de complément et les fonctions personnalisées, de s’exécuter dans le même runtime et de continuer à s’exécuter même lorsque le volet Office est fermé. Pour plus d’informations, consultez le [runtime partagé](../testing/runtimes.md#shared-runtime) et [les conseils relatifs à l’utilisation du runtime partagé dans votre complément Office](https://devblogs.microsoft.com/microsoft365dev/tips-for-using-the-shared-javascript-runtime-in-your-office-add-in%e2%80%af/) .

Voir aussi : [runtime de fonctions personnalisées](#custom-functions-runtime), [runtime](#runtime).

## <a name="subscription"></a>Abonnement

**L’abonnement** fait référence aux versions d’Office disponibles avec un abonnement Microsoft 365.

## <a name="task-pane"></a>volet Office

Les volets Office sont des surfaces d’interface ou des vues web qui apparaissent généralement sur le côté droit de la fenêtre dans Excel, Outlook, PowerPoint et Word. Les volets des tâches permettent aux utilisateurs d’accéder aux contrôles d’interface qui exécutent le code pour modifier des documents ou des e-mails, ou afficher des données d’une source de données. Utilisez les volets Office lorsque vous n’avez pas besoin ou ne pouvez pas incorporer de fonctionnalités directement dans le document. Pour en savoir plus, consultez [les volets Office dans les compléments Office](../design/task-pane-add-ins.md) .

Voir aussi : [webview](#webview).

## <a name="tutorial"></a>Tutoriel

Un **didacticiel** est une aide pédagogique conçue pour aider les gens à apprendre à utiliser un produit ou une procédure. Dans le contexte des compléments Office, un didacticiel guide un développeur de compléments tout au long du processus de développement de complément complet pour une application particulière, telle qu’Excel. Cela implique de suivre 20 étapes ou plus et représente un investissement de temps plus important qu’un [démarrage rapide](#quick-start).

Voir aussi : [démarrage rapide](#quick-start).

## <a name="volume-licensed-volume-licensed-perpetual-volume-licensing"></a>licence en volume, licence en volume perpétuelle, licence en volume

**La licence en volume** fait référence à une version perpétuelle d’Office disponible par le biais d’un contrat de licence en volume entre Microsoft et votre entreprise.

D’autres contenus Microsoft peuvent utiliser le terme **commercial** pour représenter ce concept.

Voir aussi : [perpétuel](#perpetual)

## <a name="web-add-in"></a>complément web

**Le complément web** est un terme hérité pour un complément Office. Ce terme peut être utilisé lorsque la documentation Microsoft 365 doit distinguer les compléments Office modernes des autres types de compléments tels que VBA, COM ou VSTO.

Voir aussi : [complément](#add-in).

## <a name="webview"></a>Webview

Une **vue web** est un élément ou une vue qui affiche du contenu web à l’intérieur d’une application. Les compléments de contenu et les volets Office contiennent tous deux des navigateurs web incorporés et sont des exemples de vues web dans les compléments Office.

Voir aussi : [complément de contenu](#content-add-in), [volet Office](#task-pane).

## <a name="xll"></a>XLL

Un complément **XLL** est un fichier de complément Excel qui fournit des fonctions définies par l’utilisateur et possède l’extension de fichier **.xll**. Un fichier XLL est un type de fichier de bibliothèque de liens dynamiques (DLL) qui ne peut être ouvert que par Excel. Les fichiers de complément XLL doivent être écrits en C ou C++. Les fonctions personnalisées sont l’équivalent moderne des fonctions XLL définies par l’utilisateur. Les fonctions personnalisées offrent une prise en charge sur plusieurs plateformes et sont rétrocompatibles avec les fichiers XLL. Pour plus d’informations, consultez [Étendre des fonctions personnalisées avec des fonctions XLL définies par l’utilisateur](/office/dev/add-ins/excel/make-custom-functions-compatible-with-xll-udf) .

Voir aussi : [fonction personnalisée](#custom-function).

## <a name="yeoman-generator-yo-office"></a>Générateur Yeoman, yo office

Le [générateur Yeoman pour les compléments Office](../develop/yeoman-generator-overview.md) utilise l’outil open source [Yeoman](https://github.com/yeoman/yo) pour générer un complément Office via la ligne de commande. `yo office` est la commande qui exécute le générateur Yeoman pour les compléments Office. Les guides de démarrage rapide des compléments Office et les didacticiels utilisent le générateur Yeoman.

## <a name="see-also"></a>Voir aussi

- [Ressources supplémentaires sur les compléments Office](resources-links-help.md)