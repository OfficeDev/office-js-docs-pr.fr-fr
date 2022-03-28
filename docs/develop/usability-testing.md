---
title: Tests d’utilisation pour les compléments Office
description: Découvrez comment tester la conception de votre add-in avec des utilisateurs réels.
ms.date: 07/08/2021
ms.localizationpriority: medium
ms.openlocfilehash: fbdb8f7474412e95aba03d2a409ac2e5bc4b25a6
ms.sourcegitcommit: b66ba72aee8ccb2916cd6012e66316df2130f640
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/26/2022
ms.locfileid: "64483511"
---
# <a name="usability-testing-for-office-add-ins"></a>Tests d’utilisation pour les compléments Office

Une excellente conception de complément tient compte des comportements des utilisateurs. Étant donné que vos propres idées préconçues influenceront vos décisions de conception, il est important de tester les conceptions avec des utilisateurs réels pour vous assurer que vos compléments fonctionnent correctement pour vos clients.

Vous pouvez exécuter les tests d’utilisation de différentes façons. Pour de nombreux développeurs de compléments, les études d’utilisation à distance sans modérateur sont les plus rentables et les plus rapides. Plusieurs services de test populaires le rendent facile . Voici quelques exemples.

- [UserTesting.com](https://www.UserTesting.com)
- [Optimalworkshop.com](https://www.Optimalworkshop.com)
- [Userzoom.com](https://www.Userzoom.com)

Ces services de test vous aident à simplifier la création d’un plan de test et éliminent le besoin de rechercher des participants ou de modérer les tests.

Vous avez seulement besoin de cinq participants pour révéler la plupart des problèmes d’utilisation dans votre conception. Incorporez régulièrement des petits tests dans votre cycle de développement pour vous assurer que votre produit est centré sur l’utilisateur.

> [!NOTE]
> Nous vous recommandons de tester l’utilisation de votre complément sur plusieurs plateformes. Pour [publier votre complément dans AppSource](/office/dev/store/submit-to-appsource-via-partner-center), il doit fonctionner sur toutes les [plateformes qui prennent en charge les méthodes que vous définissez](/javascript/api/requirement-sets).

## <a name="1-sign-up-for-a-testing-service"></a>1. S’inscrire à un service de test

Pour plus d’informations, reportez-vous à la section sur la [sélection d’un outil en ligne pour les tests utilisateur distants sans modérateur](https://www.nngroup.com/articles/unmoderated-user-testing-tools/).

## <a name="2-develop-your-research-questions"></a>2. Développez vos questions de recherche

Les questions de recherche définissent les objectifs de votre recherche et guident votre plan de test. Vos questions vous aideront à identifier les participants à recruter et les tâches qu’ils exécuteront. Rédigez vos questions de recherche avec autant de précision que possible. Vous pouvez également rechercher des réponses à des questions plus larges.

Voici quelques exemples de questions de recherche.

**Spécifiques**

- Les utilisateurs remarquent-ils le lien indiquant « version d’évaluation gratuite » sur la page d’accueil ?
- Lorsque les utilisateurs insèrent du contenu dans leur document à partir du complément, savent-ils où il est inséré dans le document ?

**Larges**

- Quelles sont les difficultés majeurs pour l’utilisateur dans notre complément ?
- Les utilisateurs comprennent-ils la signification des icônes dans notre barre de commandes avant de cliquer dessus ?
- Le menu des paramètres est-il facilement accessible pour les utilisateurs ?

Il est important d’obtenir des données sur l’intégralité du parcours des utilisateurs, de la découverte de votre complément jusqu’à son installation et son utilisation. Prenez en compte les questions de recherche qui abordent les aspects suivants de l’expérience utilisateur du add-in.

- Recherche de votre complément dans AppSource
- Décision d’installation de votre complément
- Première exécution
- Commandes du ruban
- Interface utilisateur du complément
- Interaction du complément avec l’espace dédié aux documents de l’application Office
- Niveau de contrôle de l’utilisateur sur les flux d’insertion de contenu

Pour plus d’informations, reportez-vous à la rubrique relative à la [collecte des réponses factuelles et des données subjectives](https://help.usertesting.com/hc/articles/115003378572-Writing-effective-questions).

## <a name="3-identify-participants-to-target"></a>3. Identifiez les participants à cibler

Les services de test à distance peuvent permettre de contrôler de nombreuses caractéristiques de vos participants aux tests. Réfléchissez soigneusement aux types d’utilisateurs que vous souhaitez cibler. Lors des premières étapes de collecte de données, il peut être préférable de recruter un large éventail de participants différents pour identifier les problèmes d’utilisation plus évidents. Plus tard, vous pourrez décider de cibler des groupes comme les utilisateurs Office avancés, des professions en particulier ou des tranches d’âge spécifiques.

## <a name="4-create-the-participant-screener"></a>4. Créez le filtre de participants

Le filtre est l’ensemble de questions et d’exigences que vous allez présenter aux participants potentiels des tests afin de les filtrer. N’oubliez pas que les participants pour les services tels que UserTesting.com sont financièrement motivés pour être sélectionnés pour votre test. Il est conseillé d’inclure des questions piège dans votre filtre si vous souhaitez exclure certains utilisateurs des tests. 

Par exemple, si vous recherchez des participants qui connaissent GitHub, pour exclure les utilisateurs qui mentent, incluez des réponses qui ne vous intéressent pas dans la liste des réponses possibles.

**Parmi les référentiels de code source suivants, lesquels connaissez-vous ?**  
 a. SourceShelf  [*Réponse disqualifiante*]  
 b. CodeContainer  [*Réponse disqualifiante*]  
 c. GitHub  [*Doit sélectionner cette réponse*]  
 d. BitBucket  [*Peut sélectionner cette réponse*]  
 e. CloudForge  [*Peut sélectionner cette réponse*]  

Si vous envisagez de tester une version déjà en ligne de votre complément, les questions suivantes peuvent permettre de sélectionner les utilisateurs qui seront en mesure de le faire.

**Pour ce test, vous devez avoir la dernière version de Microsoft PowerPoint. Avez-vous la dernière version de PowerPoint ?**  
 a. Oui [*Doit sélectionner cette réponse*]  
 b. Non [*Réponse disqualifiante*]  
 c. Je ne sais pas [*Réponse disqualifiante*]  

**Pour ce test, vous devez installer un complément gratuit pour PowerPoint et créer un compte gratuit pour l’utiliser. Êtes-vous prêt à installer un complément et à créer un compte gratuit ?**  
 a. Oui [*Doit sélectionner cette réponse*]  
 b. Non [*Réponse disqualifiante*]  

Pour plus d’informations, consultez les [meilleures pratiques en matière de questions de filtrage](https://help.usertesting.com/hc/articles/115003370731-Screener-question-best-practices).

## <a name="5-create-tasks-and-questions-for-participants"></a>5. Créez des tâches et des questions pour les participants

Essayez de classer ce que vous voulez tester par ordre de priorité afin de limiter le nombre de tâches et de questions pour le participant. Certains services paient les participants uniquement pour une certaine durée ; veillez donc à ne pas dépasser ce temps.

Autant que possible, essayez d’observer les comportements des participants au lieu de les interroger sur leurs comportements. Si vous avez besoin les interroger sur leurs comportements, demandez-leur ce qu’ils ont déjà fait, plutôt que ce qu’ils feraient dans telle ou telle situation. De cette façon, les résultats seront plus fiables.

La principale difficulté lors des tests sans modérateur consiste à s’assurer que vos participants comprennent vos tâches et vos scénarios. Vos instructions doivent être *claires et concises*. Inévitablement, si vos instructions ne sont pas claires, certains participants ne les comprendront pas.

Ne supposez jamais que l’utilisateur sera sur l’écran où il est censé être pendant le test. Vous pouvez lui indiquer l’écran sur lequel il doit se trouver au début de la tâche suivante.

Pour plus d’informations, consultez la section expliquant [comment rédiger des instructions efficaces pour les tâches](https://help.usertesting.com/hc/articles/115003371651-Writing-great-tasks).

## <a name="6-create-a-prototype-to-match-the-tasks-and-questions"></a>6. Créez un prototype pour faire correspondre les tâches et les questions

Vous pouvez faire tester votre complément déjà en ligne ou vous pouvez faire tester un prototype. N’oubliez pas que si vous souhaitez tester un complément déjà en ligne, vous devez filtrer les participants pour ne sélectionner que ceux qui ont la dernière version d’Office, qui sont prêts à installer le complément et qui sont prêts à créer un compte (sauf si vous leur fournissez des informations d’identification). Vous devez ensuite vous assurer qu’ils installent correctement votre complément.

En moyenne, aider un utilisateur à installer un complément prend environ 5 minutes. Voici un exemple d’étapes d’installation claires et concises. Modifiez-les en fonction des caractéristiques propres à votre test.

**Installez le module (insérez le nom de votre PowerPoint), en suivant les instructions ci-après.**

1. Ouvrez Microsoft PowerPoint.
1. Sélectionnez **Nouvelle présentation**.
1. Accédez à **Insérer > Mes compléments**.
1. Dans la fenêtre contextuelle, choisissez **Magasin**.
1. Saisissez (nom du complément) dans la zone de recherche.
1. Choisissez (nom du complément).
1. Prenez quelques instants pour examiner la page du magasin et pour vous familiariser avec le complément.
1. Choisissez **Ajouter** pour installer le complément.

Vous pouvez tester un prototype à n’importe quel niveau d’interaction et de fidélité visuelle. Pour une liaison et une interactivité plus complexes, pensez à utiliser un outil de prototypage tel que [InVision](https://www.invisionapp.com). Si vous souhaitez simplement tester des écrans statiques, vous pouvez héberger les images en ligne et envoyer l’URL correspondante aux participants, ou leur donner un lien vers une présentation PowerPoint en ligne. 

## <a name="7-run-a-pilot-test"></a>7. Effectuez un test pilote

Il peut être difficile de mettre au point le prototype approprié et la liste de tâches/question adéquate. Les utilisateurs peuvent ne pas comprendre certaines tâches, ou se perdre dans votre prototype. Vous devez exécuter un test pilote avec 1 à 3 utilisateurs pour solutionner les problèmes inévitables au niveau du format du test. Cette opération permet de s’assurer que vos questions sont claires, que le prototype est correctement configuré et que vous allez pouvoir recueillir le type de données que vous recherchez.

## <a name="8-run-the-test"></a>8. Lancez le test

Une fois que vous avez commandé votre test, vous obtenez des notifications par courrier électronique lorsque les participants l’effectuent. Sauf si vous avez ciblé un groupe spécifique de participants, les tests sont généralement effectués en quelques heures.

## <a name="9-analyze-results"></a>9. Analysez les résultats

Vous devez maintenant essayer d’interpréter les données que vous avez collectées. Pendant que vous regardez les vidéos des tests, notez les problèmes que rencontre l’utilisateur, ainsi que les points positifs. N’essayez pas d’interpréter la signification des données tant que vous n’avez pas affiché tous les résultats.

Un problème d’utilisation rencontré par un seul participant n’est pas suffisant pour justifier une modification de la conception. Deux ou plusieurs participants rencontrant le même problème suggère que d’autres utilisateurs dans la population globale rencontreront également ce problème.

En règle générale, soyez prudent lorsque vous utilisez vos données pour tirer des conclusions. N’essayez pas d’interpréter les données de sorte qu’elles aillent dans un sens en particulier. Ne tombez pas dans ce piège. Soyez honnête lorsque vous identifiez ce que les données prouvent réellement ou ne prouvent pas, et n’hésitez pas à reconnaître que, parfois, elles ne procurent aucune information exploitable. Gardez l’esprit ouvert. Les comportements des utilisateurs vont souvent à l’encontre des attentes du concepteur.

## <a name="see-also"></a>Voir aussi

- [Réalisation de tests d’utilisation](https://whatpixel.com/howto-conduct-usability-testing/)  
- [Meilleures pratiques pour les tests d’utilisation](https://help.usertesting.com/hc/en-us/articles/115003370231-Best-practices-for-UserTesting)  
- [Réduction de la subjectivité](https://downloads.usertesting.com/white_papers/TipSheet_MinimizingBias.pdf)  
