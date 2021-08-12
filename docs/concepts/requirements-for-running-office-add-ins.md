---
title: Configuration requise pour exécuter des compléments Office
description: Découvrez les exigences en matière de client et de serveur dont un utilisateur final a besoin pour exécuter des Office de recherche.
ms.date: 07/27/2021
localization_priority: Normal
ms.openlocfilehash: 1cc591db443c1fb0e2ca934cd05f52ad41ed61cf977ef4053af70d536867a6db
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/07/2021
ms.locfileid: "57082987"
---
# <a name="requirements-for-running-office-add-ins"></a>Configuration requise pour exécuter des compléments Office

Cet article décrit la configuration logicielle et matérielle requise pour l’exécution des compléments Office.

[!INCLUDE [publish policies note](../includes/note-publish-policies.md)]

Pour obtenir une vue d’Office les applications Office sont actuellement pris en charge, voir Office application cliente et disponibilité de plateforme pour les Office de [recherche.](../overview/office-add-in-availability.md)

## <a name="server-requirements"></a>Exigences en matière de serveur

Pour pouvoir installer et exécuter des Complément Office, vous devez d’abord déployer les fichiers manifeste et de pages web pour l’interface utilisateur et le code de votre complément sur les emplacements de serveur appropriés.

Pour tous les types de complément (compléments de contenu, Outlook et volet Office, et les commandes de compléments), vous devez déployer les fichiers de pages web de votre complément sur un serveur web ou un service d’hébergement web, tel que [Microsoft Azure](../publish/host-an-office-add-in-on-microsoft-azure.md).

[!include[HTTPS guidance](../includes/https-guidance.md)]

> [!TIP]
> lorsque vous développez et déboguez un complément dans Visual Studio, Visual Studio déploie et exécute les fichiers de page web de votre complément localement avec IIS Express et ne nécessite aucun serveur web supplémentaires.

Pour les applications clientes de contenu et du volet Des tâches, dans les applications clientes Office pris en charge (Excel, PowerPoint, Project ou Word), vous avez également besoin d’un catalogue d’applications sur SharePoint pour télécharger le fichier manifeste XML du module, ou vous devez déployer le module à l’aide des applications [](../publish/publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md) intégrées. [](/microsoft-365/admin/manage/test-and-deploy-microsoft-365-apps)

Pour tester et exécuter un Outlook, le compte de messagerie Outlook de l’utilisateur doit résider sur Exchange 2013 ou une ultérieure, qui est disponible via Microsoft 365, Exchange Online ou via une installation sur site. L’utilisateur ou l’administrateur installe les fichiers manifeste pour les compléments Outlook sur ce serveur.

> [!NOTE]
> Les comptes de messagerie POP et IMAP dans Outlook ne prennent pas en charge les compléments Office.

## <a name="client-requirements-windows-desktop-and-tablet"></a>Exigences en matière de client : ordinateur de bureau et tablette Windows

Le logiciel suivant est requis pour le développement d’un Office pour les clients de bureau ou les clients web Office pris en charge qui s’exécutent sur un ordinateur de bureau, un ordinateur portable ou une tablette Windows.

- Pour les ordinateurs de bureau Windows x86 et x64 et les tablettes telles que Surface Pro :
  - La version 32 bits ou 64 bits d’Office 2013 ou une version ultérieure s’exécutant sur Windows 7 ou une version ultérieure.
  - Excel 2013, Outlook 2013, PowerPoint 2013, Project Professionnel 2013, Project 2013 SP1 ou Word 2013, ou une version ultérieure du client Office, si vous testez ou exécutez un Complément Office, notamment pour l’un de ces clients de bureau Office. Les clients de bureau Office peuvent être installés sur site ou par le biais de « Démarrer en un clic » sur l’ordinateur client.

  Si vous avez un abonnement Microsoft 365 valide et que vous n’avez pas accès au client Office, vous pouvez télécharger et installer la dernière version de [Office](https://support.office.com/article/download-and-install-or-reinstall-office-365-or-office-2019-on-a-pc-or-mac-4414eaaf-0478-48be-9c42-23adc4716658).

- Internet Explorer 11 ou Microsoft Edge (selon les versions de Windows et Office) doit être installé, sans être nécessairement le navigateur par défaut. Pour prendre en charge les compléments Office, le client Office servant d’hôte utilise des composants de navigateur faisant partie d’Internet Explorer 11 ou de Microsoft Edge. Pour plus d’informations, voir [Navigateurs utilisés par les compléments Office](browsers-used-by-office-web-add-ins.md).

  > [!NOTE]
  > La Configuration de sécurité renforcée d’Internet Explorer (ESC) doit être désactivée pour que les compléments web Office fonctionnent. Si vous utilisez un ordinateur Windows Server comme votre client lors du développement des compléments, notez qu’ESC est activée par défaut dans Windows Server.

- L’un des éléments suivants en tant que navigateur par défaut : Internet Explorer 11 ou la dernière version de Microsoft Edge, Chrome, Firefox ou Safari (Mac OS).
- Un éditeur HTML et JavaScript tel que le Bloc-notes, [Visual Studio et les outils de développement Office ](https://www.visualstudio.com/features/office-tools-vs) ou un outil de développement web tiers.

## <a name="client-requirements-os-x-desktop"></a>Exigences en matière de client : ordinateur de bureau OS X

Outlook sur Mac, qui est distribué dans le cadre de Microsoft 365, prend en charge Outlook des macros. L’exécution de Outlook dans Outlook sur Mac a les mêmes exigences que Outlook sur Mac lui-même : le système d’exploitation doit être au moins OS X v10.10 « Yosemite ». Comme Outlook sur Mac utilise WebKit comme moteur de disposition pour restituer les pages de complément, il n’existe pas de dépendance de navigateur supplémentaire.

Les versions de client minimales d’Office pour Mac prenant en charge les compléments Office sont les suivantes.

- Version Word 15.18 (160109)
- Version Excel 15.19 (160206)
- Version PowerPoint 15.24 (160614)

## <a name="client-requirements-browser-support-for-office-web-clients-and-sharepoint"></a>Exigences en matière de client : prise en charge du navigateur pour les clients web Office et SharePoint

Tout navigateur prenant en charge ECMAScript 5.1, HTML5 et CSS3, tel qu’Internet Explorer 11 ou la dernière version de Microsoft Edge, Chrome, Firefox ou Safari (Mac OS).


## <a name="client-requirements-non-windows-smartphone-and-tablet"></a>Exigences en matière de client : smartphones et tablettes autres que Windows

L’installation du logiciel suivant est nécessaire pour tester et exécuter les compléments Outlook. Ce logiciel est conçu spécialement pour les versions d’Outlook exécutées dans des navigateurs sur smartphones et tablettes utilisant d’autres systèmes d’exploitation que Windows.


| Application Office | Appareil | Système d’exploitation | Compte Exchange | Navigateur mobile |
|:-----|:-----|:-----|:-----|:-----|
|Outlook sur Android|Tablettes et smartphones Android|Android KitKat 4.4 et version ultérieure|Sur la dernière mise à jour Applications Microsoft 365 pour les PME ou Exchange Online|Application native pour Android, navigateur non applicable|
|Outlook sur iOS|Tablettes iPad, smartphones iPhone|iOS 11 ou version ultérieure|Sur la dernière mise à jour Applications Microsoft 365 pour les PME ou Exchange Online|Application native pour iOS, navigateur non applicable|
|Outlook sur le web|iPhone 4, iPad 2, iPod Touch 4 (ou version ultérieure de ces appareils)|iOS 5 ou version ultérieure|Sur Microsoft 365, Exchange Online ou en local sur Exchange Server 2013 ou ultérieure|Safari|

> [!NOTE]
> Les applications natives OWA pour Android, OWA pour iPad et OWA pour iPhone ont été [supprimées](https://support.office.com/article/Microsoft-OWA-mobile-apps-are-being-retired-076ec122-4576-4900-bc26-937f84d25a4b) et ne sont plus requises ou disponibles pour les tests des compléments Outlook.


## <a name="see-also"></a>Voir aussi

- [Vue d’ensemble de la plateforme des compléments Office](../overview/office-add-ins.md)
- [Application cliente Office et disponibilité de la plateforme pour les compléments Office](../overview/office-add-in-availability.md)
- [Navigateurs utilisés par les compléments Office](browsers-used-by-office-web-add-ins.md)
