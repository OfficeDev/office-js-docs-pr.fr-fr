---
title: Visionneuses web utilisées par les compléments Office
description: ''
ms.date: 05/03/2019
localization_priority: Priority
ms.openlocfilehash: 6cb0d6e97dd559727b6a1e140d8417e1146e479a
ms.sourcegitcommit: 944cbb5c6ce055f6db1833182b24d490d1dce01d
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/14/2019
ms.locfileid: "33992125"
---
# <a name="web-viewers-used-by-office-add-ins"></a><span data-ttu-id="a0b06-102">Visionneuses web utilisées par les compléments Office</span><span class="sxs-lookup"><span data-stu-id="a0b06-102">Web viewers used by Office Add-ins</span></span>

<span data-ttu-id="a0b06-103">Les compléments Office étant des applications web, ils ont besoin d’une visionneuse web pour afficher les pages HTML de l’application web et d’un moteur JavaScript pour exécuter le code JavaScript.</span><span class="sxs-lookup"><span data-stu-id="a0b06-103">Since Office Add-ins are web applications, they need a web page viewer to display the HTML pages of the web application and a JavaScript engine to run the JavaScript.</span></span> <span data-ttu-id="a0b06-104">Ces deux éléments sont fournis par un navigateur installé sur l’ordinateur de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="a0b06-104">Both are supplied by a browser installed on the user’s computer.</span></span>

<span data-ttu-id="a0b06-105">Le navigateur utilisé dépend de ce qui suit :</span><span class="sxs-lookup"><span data-stu-id="a0b06-105">Which browser is used depends on:</span></span>

- <span data-ttu-id="a0b06-106">Système d’exploitation de l’ordinateur.</span><span class="sxs-lookup"><span data-stu-id="a0b06-106">The computer’s operating system.</span></span>
- <span data-ttu-id="a0b06-107">Exécution du complément dans Office Online, Office 365, Office 2013 sans abonnement ou version ultérieure.</span><span class="sxs-lookup"><span data-stu-id="a0b06-107">Whether the add-in is running in Office Online, Office 365, or non-subscription Office 2013 or later.</span></span>

<span data-ttu-id="a0b06-108">Le tableau ci-dessous répertorie le navigateur utilisé selon les plateformes et systèmes d’exploitation.</span><span class="sxs-lookup"><span data-stu-id="a0b06-108">The following table shows which browser is used for the various platforms and operating systems.</span></span>

|<span data-ttu-id="a0b06-109">**Système d’exploitation/Plateforme**</span><span class="sxs-lookup"><span data-stu-id="a0b06-109">**OS / Platform**</span></span>|<span data-ttu-id="a0b06-110">**Navigateur**</span><span class="sxs-lookup"><span data-stu-id="a0b06-110">**Browser**</span></span>|
|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
|<span data-ttu-id="a0b06-111">Office Online</span><span class="sxs-lookup"><span data-stu-id="a0b06-111">Office Online</span></span>|<span data-ttu-id="a0b06-112">Navigateur dans lequel Office Online est ouvert.</span><span class="sxs-lookup"><span data-stu-id="a0b06-112">The browser in which Office Online is opened.</span></span>|
|<span data-ttu-id="a0b06-113">Mac</span><span class="sxs-lookup"><span data-stu-id="a0b06-113">Mac</span></span>|<span data-ttu-id="a0b06-114">Safari</span><span class="sxs-lookup"><span data-stu-id="a0b06-114">Safari</span></span>|
|<span data-ttu-id="a0b06-115">iOS</span><span class="sxs-lookup"><span data-stu-id="a0b06-115">iOS</span></span>|<span data-ttu-id="a0b06-116">Safari</span><span class="sxs-lookup"><span data-stu-id="a0b06-116">Safari</span></span>|
|<span data-ttu-id="a0b06-117">Android</span><span class="sxs-lookup"><span data-stu-id="a0b06-117">Android</span></span>|<span data-ttu-id="a0b06-118">Chrome</span><span class="sxs-lookup"><span data-stu-id="a0b06-118">Chrome</span></span>|
|<span data-ttu-id="a0b06-119">Windows/Office 2013 sans abonnement ou version ultérieure</span><span class="sxs-lookup"><span data-stu-id="a0b06-119">Windows / non-subscription Office 2013 or later</span></span>|<span data-ttu-id="a0b06-120">Internet Explorer 11</span><span class="sxs-lookup"><span data-stu-id="a0b06-120">Internet Explorer 11</span></span>|
|<span data-ttu-id="a0b06-121">Windows 10 version</span><span class="sxs-lookup"><span data-stu-id="a0b06-121">Windows 10 ver.</span></span> <span data-ttu-id="a0b06-122">< 1903/Office 365</span><span class="sxs-lookup"><span data-stu-id="a0b06-122">< 1903 / Office 365</span></span>|<span data-ttu-id="a0b06-123">Internet Explorer 11</span><span class="sxs-lookup"><span data-stu-id="a0b06-123">Internet Explorer 11</span></span>|
|<span data-ttu-id="a0b06-124">Windows 10 version</span><span class="sxs-lookup"><span data-stu-id="a0b06-124">Windows 10 ver.</span></span> <span data-ttu-id="a0b06-125">>= 1903/Office 365 version < 16.0.11629</span><span class="sxs-lookup"><span data-stu-id="a0b06-125">>= 1903 / Office 365 ver < 16.0.11629</span></span>|<span data-ttu-id="a0b06-126">Internet Explorer 11</span><span class="sxs-lookup"><span data-stu-id="a0b06-126">Internet Explorer 11</span></span>|
|<span data-ttu-id="a0b06-127">Windows 10 version</span><span class="sxs-lookup"><span data-stu-id="a0b06-127">Windows 10 ver.</span></span> <span data-ttu-id="a0b06-128">>= 1903/Office 365 version >= 16.0.11629</span><span class="sxs-lookup"><span data-stu-id="a0b06-128">>= 1903 / Office 365 ver >= 16.0.11629</span></span>|<span data-ttu-id="a0b06-129">Microsoft Edge\*</span><span class="sxs-lookup"><span data-stu-id="a0b06-129">Microsoft Edge\*</span></span>|

<span data-ttu-id="a0b06-130">\*Si Microsoft Edge est utilisé, le Narrateur Windows 10 (parfois appelé « lecteur d’écran ») lit la balise `<title>` de la page qui s’ouvre dans le volet Office.</span><span class="sxs-lookup"><span data-stu-id="a0b06-130">\* When Microsoft Edge is being used, the Windows 10 Narrator (sometimes called a "screen reader") reads the `<title>` tag in the page that opens in the task pane.</span></span> <span data-ttu-id="a0b06-131">Si Internet Explorer 11 est utilisé, le Narrateur lit la barre de titre du volet Office, qui provient de la valeur `<DisplayName>` du manifeste du complément.</span><span class="sxs-lookup"><span data-stu-id="a0b06-131">When Internet Explorer 11 is being used, the Narrator reads the title bar of the task pane, which comes from the `<DisplayName>` value in the add-in's manifest.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="a0b06-132">Internet Explorer 11 ne prend pas en charge les versions de JavaScript ultérieures à la version ES5.</span><span class="sxs-lookup"><span data-stu-id="a0b06-132">Internet Explorer 11 does not support JavaScript versions later than ES5.</span></span> <span data-ttu-id="a0b06-133">Si un des utilisateurs de votre complément dispose d’une plateforme utilisant Internet Explorer 11, vous devez transpiler JavaScript vers la version ES5 ou utiliser un polyfill pour lui permettre d’utiliser la syntaxe et les fonctionnalités d’ECMAScript 2015 ou version ultérieure.</span><span class="sxs-lookup"><span data-stu-id="a0b06-133">If any of your add-in's users have platforms that use Internet Explorer 11, then to use the syntax and features of ECMAScript 2015 or later, you will need to either transpile your JavaScript to ES5 or use a polyfill.</span></span> <span data-ttu-id="a0b06-134">Par ailleurs, Internet Explorer 11 ne prend pas en charge certaines fonctionnalités HTML5 telles que les éléments multimédias, l’enregistrement et l’emplacement.</span><span class="sxs-lookup"><span data-stu-id="a0b06-134">Also, Internet Explorer 11 does not support some HTML5 features such as media, recording, and location.</span></span>

> [!NOTE]
> <span data-ttu-id="a0b06-135">En attendant leur mise à la disposition générale, vous devez participer au programme Windows Insider pour obtenir Windows 1903 ou version ultérieure, ainsi qu’au programme Office Insider pour obtenir la version 16.0.11629 ou ultérieure.</span><span class="sxs-lookup"><span data-stu-id="a0b06-135">Until they are generally available, you need to be a Windows Insider to get a Windows version 1903 or greater, and you need to be an Office Insider to get Office version 16.0.11629 or greater.</span></span>
>
> <span data-ttu-id="a0b06-136">Pour participer au programme Windows Insider :</span><span class="sxs-lookup"><span data-stu-id="a0b06-136">To join Windows Insiders:</span></span>
> 
> 1. <span data-ttu-id="a0b06-137">Accédez à [Windows Insider](https://insider.windows.com) et cliquez sur le lien pour participer au programme Windows Insider.</span><span class="sxs-lookup"><span data-stu-id="a0b06-137">Go to [Windows Insider](https://insider.windows.com) and click the link to join Windows Insiders.</span></span>
> 2. <span data-ttu-id="a0b06-138">Vous accédez alors à une page d’instructions sur l’utilisation des paramètres Windows pour activer les builds Windows.</span><span class="sxs-lookup"><span data-stu-id="a0b06-138">You will be taken to a page with instructions about how to use Windows Settings to enable preview builds of Windows.</span></span> <span data-ttu-id="a0b06-139">Suivez les instructions.</span><span class="sxs-lookup"><span data-stu-id="a0b06-139">Follow the instructions on the page.</span></span> <span data-ttu-id="a0b06-140">Lorsque vous sélectionnez le rythme des mises à jour, choisissez l’option la plus rapide.</span><span class="sxs-lookup"><span data-stu-id="a0b06-140">When you select the pace of updates, choose the fastest option.</span></span>
>
> <span data-ttu-id="a0b06-141">Pour participer au programme Office Insider :</span><span class="sxs-lookup"><span data-stu-id="a0b06-141">To join Office Insiders:</span></span>
> 
> 1. <span data-ttu-id="a0b06-142">Accédez à [Participer au programme Office Insider](https://insider.office.com/join).</span><span class="sxs-lookup"><span data-stu-id="a0b06-142">Go to [Get started as an Office Insider](https://insider.office.com/join).</span></span>
> 2. <span data-ttu-id="a0b06-143">Suivez les instructions détaillées sur cette page.</span><span class="sxs-lookup"><span data-stu-id="a0b06-143">Follow the instruction on that page to join.</span></span> <span data-ttu-id="a0b06-144">Lorsque vous êtes invité à spécifier un canal, sélectionnez Insider.</span><span class="sxs-lookup"><span data-stu-id="a0b06-144">When asked to specify a channel, select Insider.</span></span>

## <a name="see-also"></a><span data-ttu-id="a0b06-145">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="a0b06-145">See also</span></span>

- [<span data-ttu-id="a0b06-146">Configuration requise pour exécuter des compléments Office</span><span class="sxs-lookup"><span data-stu-id="a0b06-146">Requirements for running Office Add-ins</span></span>](requirements-for-running-office-add-ins.md)
