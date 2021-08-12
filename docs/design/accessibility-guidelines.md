---
title: Instructions d’accessibilité pour les compléments Office
description: Découvrez comment rendre votre Office de service accessible à tous les utilisateurs.
ms.date: 09/24/2018
localization_priority: Normal
ms.openlocfilehash: e9531bfcf4ffb1bca178563785a31fdc3b9e21e1544fb74e2495b8e1c51de9c2
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/07/2021
ms.locfileid: "57082743"
---
# <a name="accessibility-guidelines"></a>Conseils sur l’accessibilité

Lorsque vous concevez et développez des compléments Office, il est important de faire en sorte que tous les utilisateurs et clients potentiels puissent utiliser votre complément. Appliquez les instructions suivantes pour vous assurer que votre solution est accessible à tous les publics.

## <a name="design-for-multiple-input-methods"></a>Tenez compte des différentes méthodes d’entrée

- Veillez à ce que les utilisateurs puissent effectuer des opérations à l’aide du clavier uniquement. Les utilisateurs doivent pouvoir accéder à tous les éléments exploitables de la page en utilisant une combinaison de la touche Tab et des flèches.
- Sur un appareil mobile, lorsque les utilisateurs actionnent un contrôle en mode tactile, l’appareil doit fournir des commentaires audio utiles.
- Prévoyez des étiquettes d’aide pour tous les contrôles interactifs. 

## <a name="make-your-add-in-easy-to-use"></a>Faites en sorte que votre complément soit facile à utiliser

- Ne vous contentez pas d’utiliser un seul attribut (comme la couleur, la taille, la forme, l’emplacement, l’orientation ou le son) pour assurer la lisibilité de votre interface utilisateur.
- Évitez les changements de contexte inattendus, par exemple un déplacement de la sélection sur un autre élément de l’interface sans action de l’utilisateur.
- Fournissez un moyen de vérifier, de confirmer ou d’annuler toutes les actions qui engagent la responsabilité ou le consentement de l’utilisateur.
- Fournissez un moyen de suspendre ou d’arrêter les contenus multimédias, tels que les ressources audio et vidéo.
- N’imposez pas de limite de temps pour les actions de l’utilisateur.

## <a name="make-your-add-in-easy-to-see"></a>Améliorez la lisibilité de votre complément

- Évitez les changements de couleur inattendus.
- Fournissez des informations compréhensibles et pertinentes pour décrire les éléments de l’interface utilisateur, les titres et en-têtes, les entrées et les erreurs. Vérifiez que le nom des contrôles en décrit bien l’utilité.
- Suivez les [instructions standard](https://www.w3.org/TR/UNDERSTANDING-WCAG20/visual-audio-contrast-contrast.html) pour le contraste des couleurs.

## <a name="account-for-assistive-technologies"></a>Tenez compte des technologies d’assistance

- Évitez d’utiliser des fonctionnalités qui interfèrent avec les technologies d’assistance, notamment en ce qui concerne les interactions visuelles, audio ou autres.
- Ne fournissez pas de texte dans un format image. Les lecteurs d’écran ne peuvent pas lire le texte dans les images.
- Fournissez un moyen aux utilisateurs d’ajuster ou de désactiver le son de toutes les sources audio.
- Fournissez un moyen aux utilisateurs d’activer des légendes ou une description audio avec les sources audio.
- Prévoyez d’autres solutions que des signaux audio pour informer les utilisateurs, telles que des indications visuelles ou des vibrations.

## <a name="see-also"></a>Voir aussi

- [Web Content Accessibility Guidelines (WCAG) 2.0](https://www.w3.org/TR/wcag2ict/#REF-WCAG20)
- [Conseils sur l’application de WCAG 2.0 aux technologies d’information et de communication non web (WCAG2ICT)](https://www.w3.org/TR/wcag2ict/)
- [Norme européenne sur les conditions d’accessibilité pour les technologies d’information et de communication (ICT)](https://www.etsi.org/deliver/etsi_en/301500_301599/301549/01.00.00_20/en_301549v010000c.pdf) 
