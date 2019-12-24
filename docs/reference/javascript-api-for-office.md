---
layout: LandingPage
ms.topic: landing-page
title: Documentation référence de l’API JavaScript pour Office
description: En savoir plus sur les API JavaScript pour Office.
ms.date: 12/24/2019
localization_priority: Priority
ms.openlocfilehash: c10eeb5c89a74b28e9af44bf72b20a7ad610738b
ms.sourcegitcommit: 350f5c6954dec3e9384e2030cd3265aaba7ae904
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 12/23/2019
ms.locfileid: "40851550"
---
# <a name="api-reference-documentation"></a>Documentation de référence de l'API

Un complément peut utiliser les API JavaScript pour Office pour interagir avec des objets dans les applications hôtes Office. 

<ul>
    <li><b>Les API propres aux hôtes</b> fournissent des objets fortement typés qui peuvent être utilisés pour interagir avec des objets natifs d’une application Office spécifique.</li>
    <li>Les API <b>Communes</b> peuvent être utilisées pour accéder à des fonctionnalités telles qu’une interface utilisateur, des boîtes de dialogue et des paramètres du client, qui sont communes à plusieurs types d’applications Office.</li>
</ul>

Vous devez utiliser les API propres à l’hôte dans la mesure du possible, et utiliser les API communes uniquement pour les scénarios qui ne sont pas pris en charge par les API propres à l’hôte. Si vous souhaitez en savoir plus sur ces deux modèles API, consultez<a href="../overview/office-add-ins-fundamentals.md#api-models">Création de compléments Office</a>.

<h2>Référence d’API</h2>

<ul class="panelContent cardsF cols cols3">
    <li>
        <div class="cardSize">
            <div class="cardPadding">
                <div class="card">
                    <div class="cardImageOuter">
                        <div class="cardImage">
                            <a href="/javascript/api/excel"><img src="../images/index/logo-excel.svg" alt="Excel API reference docs" /></a>
                        </div>
                    </div>
                    <div class="cardText">
                        <h3>Référence de l’API pour Excel</h3>
                        <p><a href="/javascript/api/excel">API JavaScript pour la création de compléments Excel.</a></p>
                    </div>
                </div>
            </div>
        </div>
    </li>
    <li>
        <div class="cardSize">
            <div class="cardPadding">
                <div class="card">
                    <div class="cardImageOuter">
                        <div class="cardImage">
                            <a href="/javascript/api/outlook"><img src="../images/index/logo-outlook.svg" alt="Outlook API reference docs" /></a>
                        </div>
                    </div>
                    <div class="cardText">
                        <h3>Documentation de référence de l’API Outlook</h3>
                        <p><a href="/javascript/api/outlook">API JavaScript pour la création de compléments Outlook.</a></p>
                    </div>
                </div>
            </div>
        </div>
    </li>
    <li>
        <div class="cardSize">
            <div class="cardPadding">
                <div class="card">
                    <div class="cardImageOuter">
                        <div class="cardImage">
                            <a href="/javascript/api/word"><img src="../images/index/logo-word.svg" alt="Word API reference docs" /></a>
                        </div>
                    </div>
                    <div class="cardText">
                        <h3>Référence de l’API pour Word</h3>
                        <p><a href="/javascript/api/word">API JavaScript pour la création de compléments Word.</a></p>
                    </div>
                </div>
            </div>
        </div>
    </li>
    <li>
        <div class="cardSize">
            <div class="cardPadding">
                <div class="card">
                    <div class="cardImageOuter">
                        <div class="cardImage">
                            <a href="/javascript/api/powerpoint"><img src="../images/index/logo-powerpoint.svg" alt="PowerPoint API reference docs" /></a>
                        </div>
                    </div>
                    <div class="cardText">
                        <h3>Référence API PowerPoint</h3>
                        <p><a href="/javascript/api/powerpoint">API JavaScript pour la création de compléments PowerPoint.</a></p>
                    </div>
                </div>
            </div>
        </div>
    </li>
    <li>
        <div class="cardSize">
            <div class="cardPadding">
                <div class="card">
                    <div class="cardImageOuter">
                        <div class="cardImage">
                            <a href="/javascript/api/onenote"><img src="../images/index/logo-onenote.svg" alt="OneNote API reference docs" /></a>
                        </div>
                    </div>
                    <div class="cardText">
                        <h3>Référence de l’API pour OneNote</h3>
                        <p><a href="/javascript/api/onenote">API JavaScript pour la création de compléments OneNote.</a></p>
                    </div>
                </div>
            </div>
        </div>
    </li>
    <li>
        <div class="cardSize">
            <div class="cardPadding">
                <div class="card">
                    <div class="cardImageOuter">
                        <div class="cardImage">
                            <a href="/javascript/api/office"><img src="../images/index-landing-page/i_code-blocks.svg" alt="reference docs" /></a>
                        </div>
                    </div>
                    <div class="cardText">
                        <h3>Référence d’API commune</h3>
                        <p><a href="/javascript/api/office">API JavaScript pouvant être utilisées par les compléments Office.</a></p>
                    </div>
                </div>
            </div>
        </div>
    </li>
</ul>

<b>Remarque</b>: il n’existe actuellement aucune API JavaScript propre à l’hôte pour Project ; vous utiliserez des API communes pour créer des compléments Project. de plus, l’étendue de l’API propre à l’hôte pour PowerPoint est très limitée ; vous utiliserez principalement les API communes pour créer des compléments PowerPoint.

<h2>Spécifications d’ouverture de l’API</h2>

Au fur et à mesure que nous concevons et développons de nouvelles API pour les compléments Office, nous les mettons à votre disposition sur notre page de [spécifications d’ouverture de l’API](openspec/openspec.md) pour que vous puissiez fournir vos commentaires. Découvrez les nouvelles fonctionnalités dans le pipeline et donnez votre avis sur nos spécifications de conception.