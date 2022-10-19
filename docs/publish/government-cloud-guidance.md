---
title: Conseils pour le déploiement de compléments Office sur des clouds gouvernementaux
description: Découvrez comment déployer votre complément Office dans des environnements cloud gouvernementaux sécurisés
ms.date: 10/17/2022
ms.localizationpriority: medium
ms.openlocfilehash: f3995c62a1b7fb482df6a15da870f747f55e9508
ms.sourcegitcommit: eca6c16d0bb74bed2d35a21723dd98c6b41ef507
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/18/2022
ms.locfileid: "68607590"
---
# <a name="guidance-for-deploying-office-add-ins-on-government-clouds"></a>Conseils pour le déploiement de compléments Office sur des clouds gouvernementaux

Microsoft fournit les options cloud gouvernementales pour nos clients sensibles à la confidentialité dans les organisations gouvernementales locales, d’état et nationales. Cela donne aux partenaires la possibilité de cibler des clients critiques avec leurs compléments Office. En raison de la nature plus restreinte de ces environnements, qui est importante pour les besoins de confidentialité et de sécurité des clients, toutes les ressources qui sont généralement disponibles dans un environnement de production standard ne sont pas disponibles dans ces clouds.

Pour les partenaires qui fournissent leurs compléments Office aux clients dans ces environnements cloud restreints, il existe des différences importantes par rapport à l’environnement de cloud public standard qui doivent être prises en compte. Les informations suivantes fournissent les détails qui nécessitent une gestion spéciale par les développeurs écrivant des compléments Office qui ciblent les clients dans ces environnements.

## <a name="all-sovereign-environments"></a>Tous les environnements souverains

Pour tous les environnements cloud gouvernementaux (par exemple, cloud souverain), l’Office Store public n’est pas disponible. Cela signifie que les utilisateurs finaux ne peuvent pas acquérir de compléments Office directement à partir du magasin public. Les administrateurs ne peuvent pas non plus déployer des compléments Office directement à partir du magasin public dans leur portail Administration. Au lieu de cela, vous devez travailler avec les administrateurs pour vous assurer que les éléments suivants sont les suivants :

- Les ressources et services requis pour votre solution sont disponibles à l’intérieur de la limite du cloud. Soit vous travaillez avec les administrateurs locataires pour approvisionner votre service et vos ressources à l’intérieur de la limite du cloud, soit vous travaillez avec l’administrateur réseau pour permettre l’accès à vos ressources qui résident en dehors de la limite du cloud.

- Les ressources auxquelles le complément Office accède sont conformes aux exigences du cloud public à partir duquel ils sont accessibles. Ils doivent également se conformer à toutes les exigences supplémentaires du locataire client pour lequel la solution est en cours d’approvisionnement. Ces exigences incluent le transfert, la gestion et le stockage de données potentiellement sensibles, ainsi que des procédures de contrôle d’accès et de sécurité du code et des ressources.

- Le manifeste du complément Office qui décrit la solution et son emplacement source, le cas échéant, pour le déploiement de cloud public particulier est obtenu auprès du partenaire et ingéré pour le déploiement vers les utilisateurs appropriés via le portail Administration.

## <a name="us-government-community-cloud-gcc"></a>Cloud de la communauté du gouvernement des États-Unis (GCC)

Outre les exigences applicables à tous les clouds souverains, chaque environnement de cloud souverain a ses propres considérations qui peuvent affecter les compléments Office ciblant l’environnement. GCC est le moins restrictif des environnements cloud gouvernementaux et certaines des ressources requises par le complément sont disponibles à partir de leurs points de terminaison de production habituels dans cet environnement. L’une de ces ressources est la bibliothèque d’API JavaScript Office. Les partenaires de solutions peuvent continuer à référencer la ressource Office.js publique comme ils le font avec leur solution de production publique.

## <a name="gcc-high-gcch-us-department-of-defense-dod-or-other-sovereign-managed-clouds"></a>GCC High (GCCH), US Department of Defense (DOD) ou d’autres clouds souverains gérés

Ces clouds gouvernementaux restent connectés à Internet, mais l’ensemble des points de terminaison publics mis à disposition est fortement restreint. L’un de ces points de terminaison restreints est le point de terminaison public pour le chargement de la bibliothèque d’API JavaScript Office. L’emplacement CDN public pour Office.js ne sera pas accessible à partir de ces environnements. Toutefois, il y aura un CDN Microsoft Office interne par cloud approvisionné avec la même ressource. Cela signifie que l’URL du point de terminaison pour accéder à Office.js sera différente et que votre complément Office aura peut-être besoin d’un certain niveau de personnalisation pour s’exécuter. Compte tenu des restrictions supplémentaires, il est probable que toute solution fournie aux clients nécessite également des services de fournisseur d’hébergement dans l’environnement, ce qui nécessite des personnalisations supplémentaires. Vous devez déterminer la meilleure façon de fournir votre solution aux clients, afin qu’elle soit conforme aux restrictions supplémentaires imposées aux services s’exécutant dans les limites de ces environnements.

## <a name="airgapped-sovereign-clouds"></a>Nuages souverains aérés

Ces clouds gouvernementaux sont essentiellement déconnectés de l’Internet public entièrement. Toutes les ressources qui seraient normalement accessibles à partir de ressources publiques doivent plutôt être approvisionnées sur mesure dans ces environnements cloud. Dans les clouds GCCH et DOD mentionnés précédemment, la plupart des fournisseurs de solutions (sinon tous) devront approvisionner leurs services et ressources dans le cloud. Il existe une option permettant d’effectuer des exceptions de pare-feu qui autorisent l’accès aux services et ressources publics. Toutefois, ce contournement n’est pas possible dans les nuages aériens. Comme avec les clouds GCCH et DOD, un CDN Microsoft Office est provisionné dans chaque environnement cloud qui héberge les ressources requises telles que la bibliothèque Office.js. Vous devez travailler en étroite collaboration avec les administrateurs de locataires clients pour déterminer la meilleure façon de fournir vos services et ressources d’une manière conforme aux exigences d’accès strictes pour les clouds souverains aérés.
