---
title: Instructions d’accessibilité pour les compléments Office
description: Découvrez comment faire en sorte que votre complément Office soit accessible à tous les utilisateurs.
ms.date: 09/24/2018
localization_priority: Normal
ms.openlocfilehash: 889563af8ab5f7bbcd4037eedb42933369a92cf2
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/08/2020
ms.locfileid: "44607991"
---
# <a name="accessibility-guidelines"></a><span data-ttu-id="3751c-103">Conseils sur l’accessibilité</span><span class="sxs-lookup"><span data-stu-id="3751c-103">Accessibility guidelines</span></span>

<span data-ttu-id="3751c-p101">Lorsque vous concevez et développez des compléments Office, il est important de faire en sorte que tous les utilisateurs et clients potentiels puissent utiliser votre complément. Appliquez les instructions suivantes pour vous assurer que votre solution est accessible à tous les publics.</span><span class="sxs-lookup"><span data-stu-id="3751c-p101">As you design and develop your Office Add-ins, you'll want to ensure that all potential users and customers are able to use your add-in successfully. Apply the following guidelines to ensure that your solution is accessible to all audiences.</span></span>

## <a name="design-for-multiple-input-methods"></a><span data-ttu-id="3751c-106">Tenez compte des différentes méthodes d’entrée</span><span class="sxs-lookup"><span data-stu-id="3751c-106">Design for multiple input methods</span></span>

- <span data-ttu-id="3751c-p102">Veillez à ce que les utilisateurs puissent effectuer des opérations à l’aide du clavier uniquement. Les utilisateurs doivent pouvoir accéder à tous les éléments exploitables de la page en utilisant une combinaison de la touche Tab et des flèches.</span><span class="sxs-lookup"><span data-stu-id="3751c-p102">Ensure that users can perform operations by using only the keyboard. Users should be able to move to all actionable elements on the page by using a combination of the Tab and arrow keys.</span></span>
- <span data-ttu-id="3751c-109">Sur un appareil mobile, lorsque les utilisateurs actionnent un contrôle en mode tactile, l’appareil doit fournir des commentaires audio utiles.</span><span class="sxs-lookup"><span data-stu-id="3751c-109">On a mobile device, when users operate a control by touch, the device should provide useful audio feedback.</span></span>
- <span data-ttu-id="3751c-110">Prévoyez des étiquettes d’aide pour tous les contrôles interactifs.</span><span class="sxs-lookup"><span data-stu-id="3751c-110">Provide helpful labels for all interactive controls.</span></span> 

## <a name="make-your-add-in-easy-to-use"></a><span data-ttu-id="3751c-111">Faites en sorte que votre complément soit facile à utiliser</span><span class="sxs-lookup"><span data-stu-id="3751c-111">Make your add-in easy to use</span></span>

- <span data-ttu-id="3751c-112">Ne vous contentez pas d’utiliser un seul attribut (comme la couleur, la taille, la forme, l’emplacement, l’orientation ou le son) pour assurer la lisibilité de votre interface utilisateur.</span><span class="sxs-lookup"><span data-stu-id="3751c-112">Don't rely on a single attribute, such as color, size, shape, location, orientation, or sound, to convey meaning in your UI.</span></span>
- <span data-ttu-id="3751c-113">Évitez les changements de contexte inattendus, par exemple un déplacement de la sélection sur un autre élément de l’interface sans action de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="3751c-113">Avoid unexpected changes of context, such as moving the focus to a different UI element without user action.</span></span>
- <span data-ttu-id="3751c-114">Fournissez un moyen de vérifier, de confirmer ou d’annuler toutes les actions qui engagent la responsabilité ou le consentement de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="3751c-114">Provide a way to verify, confirm, or reverse all binding actions.</span></span>
- <span data-ttu-id="3751c-115">Fournissez un moyen de suspendre ou d’arrêter les contenus multimédias, tels que les ressources audio et vidéo.</span><span class="sxs-lookup"><span data-stu-id="3751c-115">Provide a way to pause or stop media, such as audio and video.</span></span>
- <span data-ttu-id="3751c-116">N’imposez pas de limite de temps pour les actions de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="3751c-116">Do not impose a time limit for user action.</span></span>

## <a name="make-your-add-in-easy-to-see"></a><span data-ttu-id="3751c-117">Améliorez la lisibilité de votre complément</span><span class="sxs-lookup"><span data-stu-id="3751c-117">Make your add-in easy to see</span></span>

- <span data-ttu-id="3751c-118">Évitez les changements de couleur inattendus.</span><span class="sxs-lookup"><span data-stu-id="3751c-118">Avoid unexpected color changes.</span></span>
- <span data-ttu-id="3751c-p103">Fournissez des informations compréhensibles et pertinentes pour décrire les éléments de l’interface utilisateur, les titres et en-têtes, les entrées et les erreurs. Vérifiez que le nom des contrôles en décrit bien l’utilité.</span><span class="sxs-lookup"><span data-stu-id="3751c-p103">Provide meaningful and timely information to describe UI elements, titles and headings, inputs, and errors. Ensure that names of controls adequately describe the intent of the control.</span></span>
- <span data-ttu-id="3751c-121">Suivez les [instructions standard](https://www.w3.org/TR/UNDERSTANDING-WCAG20/visual-audio-contrast-contrast.html) pour le contraste des couleurs.</span><span class="sxs-lookup"><span data-stu-id="3751c-121">Follow [standard guidelines](https://www.w3.org/TR/UNDERSTANDING-WCAG20/visual-audio-contrast-contrast.html) for color contrast.</span></span>

## <a name="account-for-assistive-technologies"></a><span data-ttu-id="3751c-122">Tenez compte des technologies d’assistance</span><span class="sxs-lookup"><span data-stu-id="3751c-122">Account for assistive technologies</span></span>

- <span data-ttu-id="3751c-123">Évitez d’utiliser des fonctionnalités qui interfèrent avec les technologies d’assistance, notamment en ce qui concerne les interactions visuelles, audio ou autres.</span><span class="sxs-lookup"><span data-stu-id="3751c-123">Avoid using features that interfere with assistive technologies, including visual, audio, or other interactions.</span></span>
- <span data-ttu-id="3751c-p104">Ne fournissez pas de texte dans un format image. Les lecteurs d’écran ne peuvent pas lire le texte dans les images.</span><span class="sxs-lookup"><span data-stu-id="3751c-p104">Do not provide text in an image format. Screen readers cannot read text within images.</span></span>
- <span data-ttu-id="3751c-126">Fournissez un moyen aux utilisateurs d’ajuster ou de désactiver le son de toutes les sources audio.</span><span class="sxs-lookup"><span data-stu-id="3751c-126">Provide a way for users to adjust or mute all audio sources.</span></span>
- <span data-ttu-id="3751c-127">Fournissez un moyen aux utilisateurs d’activer des légendes ou une description audio avec les sources audio.</span><span class="sxs-lookup"><span data-stu-id="3751c-127">Provide a way for users to turn on captions or audio description with audio sources.</span></span>
- <span data-ttu-id="3751c-128">Prévoyez d’autres solutions que des signaux audio pour informer les utilisateurs, telles que des indications visuelles ou des vibrations.</span><span class="sxs-lookup"><span data-stu-id="3751c-128">Provide alternatives to sound as a means to alert users, such as visual cues or vibrations.</span></span>

## <a name="see-also"></a><span data-ttu-id="3751c-129">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="3751c-129">See also</span></span>

- [<span data-ttu-id="3751c-130">Web Content Accessibility Guidelines (WCAG) 2.0</span><span class="sxs-lookup"><span data-stu-id="3751c-130">Web Content Accessibility Guidelines (WCAG) 2.0</span></span>](https://www.w3.org/TR/wcag2ict/#REF-WCAG20)
- [<span data-ttu-id="3751c-131">Conseils sur l’application de WCAG 2.0 aux technologies d’information et de communication non web (WCAG2ICT)</span><span class="sxs-lookup"><span data-stu-id="3751c-131">Guidance on Applying WCAG 2.0 to Non-Web Information and Communications Technologies (WCAG2ICT)</span></span>](https://www.w3.org/TR/wcag2ict/)
- [<span data-ttu-id="3751c-132">Norme européenne sur les conditions d’accessibilité pour les technologies d’information et de communication (ICT)</span><span class="sxs-lookup"><span data-stu-id="3751c-132">European Standard on accessibility requirements for Information and Communication Technologies (ICT)</span></span>](https://www.etsi.org/deliver/etsi_en/301500_301599/301549/01.00.00_20/en_301549v010000c.pdf) 
