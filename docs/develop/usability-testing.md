---
title: Tests d’utilisation pour les compléments Office
description: Découvrez comment tester votre conception de complément avec des utilisateurs réels.
ms.date: 07/08/2021
ms.localizationpriority: medium
ms.openlocfilehash: 49a2af983615779160886961e8269e4588d0fc9e
ms.sourcegitcommit: 3abcf7046446e7b02679c79d9054843088312200
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 11/02/2022
ms.locfileid: "68810280"
---
# <a name="usability-testing-for-office-add-ins"></a>Tests d’utilisation pour les compléments Office

A great add-in design takes user behaviors into account. Because your own preconceptions influence your design decisions, it’s important to test designs with real users to make sure that your add-ins work well for your customers.

Vous pouvez exécuter les tests d’utilisation de différentes façons. Pour de nombreux développeurs de compléments, les études d’utilisation à distance sans modérateur sont les plus rentables et les plus rapides. Plusieurs services de test populaires facilitent cette opération; Voici quelques exemples.

- [UserTesting.com](https://www.UserTesting.com)
- [Optimalworkshop.com](https://www.Optimalworkshop.com)
- [Userzoom.com](https://www.Userzoom.com)

Ces services de test vous aident à simplifier la création d’un plan de test et éliminent le besoin de rechercher des participants ou de modérer les tests.

You need only five participants to uncover most usability issues in your design. Incorporate small tests regularly throughout your development cycle to ensure that your product is user-centered.

> [!NOTE]
> We recommend that you test the usability of your add-in across multiple platforms. To [publish your add-in to AppSource](/office/dev/store/submit-to-appsource-via-partner-center), it must work on all [platforms that support the methods that you define](/javascript/api/requirement-sets).

## <a name="1-sign-up-for-a-testing-service"></a>1. S’inscrire à un service de test

Pour plus d’informations, reportez-vous à la section sur la [sélection d’un outil en ligne pour les tests utilisateur distants sans modérateur](https://www.nngroup.com/articles/unmoderated-user-testing-tools/).

## <a name="2-develop-your-research-questions"></a>2. Développez vos questions de recherche

Research questions define the objectives of your research and guide your test plan. Your questions will help you identify participants to recruit and the tasks they will perform. Make your research questions as specific as you can. You can also seek to answer broader questions.

Voici quelques exemples de questions de recherche.

**Spécifiques**

- Les utilisateurs remarquent-ils le lien indiquant « version d’évaluation gratuite » sur la page d’accueil ?
- Lorsque les utilisateurs insèrent du contenu dans leur document à partir du complément, savent-ils où il est inséré dans le document ?

**Larges**

- Quelles sont les difficultés majeurs pour l’utilisateur dans notre complément ?
- Les utilisateurs comprennent-ils la signification des icônes dans notre barre de commandes avant de cliquer dessus ?
- Le menu des paramètres est-il facilement accessible pour les utilisateurs ?

Il est important d’obtenir des données sur l’intégralité du parcours des utilisateurs, de la découverte de votre complément jusqu’à son installation et son utilisation. Envisagez des questions de recherche qui traitent des aspects suivants de l’expérience utilisateur du complément.

- Recherche de votre complément dans AppSource
- Décision d’installation de votre complément
- Première exécution
- Commandes du ruban
- Interface utilisateur du complément
- Interaction du complément avec l’espace dédié aux documents de l’application Office
- Niveau de contrôle de l’utilisateur sur les flux d’insertion de contenu

Pour plus d’informations, reportez-vous à la rubrique relative à la [collecte des réponses factuelles et des données subjectives](https://help.usertesting.com/hc/articles/115003378572-Writing-effective-questions).

## <a name="3-identify-participants-to-target"></a>3. Identifiez les participants à cibler

Remote testing services can give you control over many characteristics of your test participants. Think carefully about what kinds of users you want to target. In your early stages of data collection, it might be better to recruit a wide variety of participants to identify more obvious usability issues. Later, you might choose to target groups like advanced Office users, particular occupations, or specific age ranges.

## <a name="4-create-the-participant-screener"></a>4. Créez le filtre de participants

The screener is the set of questions and requirements you will present to prospective test participants to screen them for your test. Keep in mind that participants for services like UserTesting.com have a financial interest in qualifying for your test. It's a good idea to include trick questions in your screener if you want to  exclude certain users from the test. 

Par exemple, si vous recherchez des participants qui connaissent GitHub, pour exclure les utilisateurs qui mentent, incluez des réponses qui ne vous intéressent pas dans la liste des réponses possibles.

**Parmi les référentiels de code source suivants, lesquels connaissez-vous ?**  
 a. SourceShelf  [*Reject*]  
 b. CodeContainer  [*Reject*]  
 c. GitHub  [*Must select*]  
 d. BitBucket  [*May select*]  
 e. CloudForge  [*May select*]  

Si vous envisagez de tester une version déjà en ligne de votre complément, les questions suivantes peuvent permettre de sélectionner les utilisateurs qui seront en mesure de le faire.

**Pour ce test, vous devez avoir la dernière version de Microsoft PowerPoint. Avez-vous la dernière version de PowerPoint ?**  
 a. Yes [*Must select*]  
 b. No [*Reject*]  
 c. I don’t know [*Reject*]  

**Pour ce test, vous devez installer un complément gratuit pour PowerPoint et créer un compte gratuit pour l’utiliser. Êtes-vous prêt à installer un complément et à créer un compte gratuit ?**  
 a. Yes [*Must select*]  
 b. No [*Reject*]  

Pour plus d’informations, consultez les [meilleures pratiques en matière de questions de filtrage](https://help.usertesting.com/hc/articles/115003370731-Screener-question-best-practices).

## <a name="5-create-tasks-and-questions-for-participants"></a>5. Créez des tâches et des questions pour les participants

Try to prioritize what you want tested so that you can limit the number of tasks and questions for the participant. Some services pay participants only for a set amount of time, so you want to make sure not to go over.

Try to observe participant behaviors instead of asking about them, whenever possible. If you need to ask about behaviors, ask about what participants have done in the past, rather than what they would expect to do in a situation. This tends to give more reliable results.

The main challenge in unmoderated testing is making sure your participants understand your tasks and scenarios. Your directions should be *clear and concise*. Inevitably, if there is potential for confusion, someone will be confused.

Don't assume that your user will be on the screen they’re supposed to be on at any given point during the test. Consider telling them what screen they need to be on to start the next task.

Pour plus d’informations, consultez la section expliquant [comment rédiger des instructions efficaces pour les tâches](https://help.usertesting.com/hc/articles/115003371651-Writing-great-tasks).

## <a name="6-create-a-prototype-to-match-the-tasks-and-questions"></a>6. Créez un prototype pour faire correspondre les tâches et les questions

Vous pouvez faire tester votre complément déjà en ligne ou vous pouvez faire tester un prototype. N’oubliez pas que si vous souhaitez tester un complément déjà en ligne, vous devez filtrer les participants pour ne sélectionner que ceux qui ont la dernière version d’Office, qui sont prêts à installer le complément et qui sont prêts à créer un compte (sauf si vous leur fournissez des informations d’identification). Vous devez ensuite vous assurer qu’ils installent correctement votre complément.

On average, it takes about 5 minutes to walk users through how to install an add-in. The following is an example of clear, concise installation steps. Adjust the steps based on the specifics of your test.

**Installez le complément (insérez le nom de votre complément ici) pour PowerPoint, en suivant les instructions suivantes.**

1. Ouvrez Microsoft PowerPoint.
1. Sélectionnez **Nouvelle présentation**.
1. Accédez à **Insérer** > **mes compléments**.
1. Dans la fenêtre contextuelle, choisissez **Stocker**.
1. Saisissez (nom du complément) dans la zone de recherche.
1. Choisissez (nom du complément).
1. Prenez quelques instants pour examiner la page du magasin et pour vous familiariser avec le complément.
1. Choisissez **Ajouter** pour installer le complément.

You can test a prototype at any level of interaction and visual fidelity. For more complex linking and interactivity, consider a prototyping tool like [InVision](https://www.invisionapp.com). If you just want to test static screens, you can host images online and send participants the corresponding URL, or give them a link to an online PowerPoint presentation. 

## <a name="7-run-a-pilot-test"></a>7. Effectuez un test pilote

It can be tricky to get the prototype and your task/question list right. Users might be confused by tasks, or might get lost in your prototype. You should run a pilot test with 1-3 users to work out the inevitable issues with the test format. This will help to ensure that your questions are clear, that the prototype is set up correctly, and that you’re capturing the type of data you’re looking for.

## <a name="8-run-the-test"></a>8. Lancez le test

After you order your test, you will get email notifications when participants complete it. Unless you’ve targeted a specific group of participants, the tests are usually completed within a few hours.

## <a name="9-analyze-results"></a>9. Analysez les résultats

This is the part where you try to make sense of the data you’ve collected. While watching the test videos, record notes about problems and successes the user has. Avoid trying to interpret the meaning of the data until you have viewed all the results.

A single participant having a usability issue is not enough to warrant making a change to the design. Two or more participants encountering the same issue suggests that other users in the general population will also encounter that issue.

In general, be careful about how you use your data to draw conclusions. Don’t fall into the trap of trying to make the data fit a certain narrative; be honest about what the data actually proves, disproves, or simply fails to provide any insight about. Keep an open mind; user behavior frequently defies designer’s expectations.

## <a name="see-also"></a>Voir aussi

- [Réalisation de tests d’utilisation](https://whatpixel.com/howto-conduct-usability-testing/)  
- [Meilleures pratiques pour les tests d’utilisation](https://help.usertesting.com/hc/articles/115003370231-Best-practices-for-UserTesting)  
- [Réduction de la subjectivité](https://downloads.usertesting.com/white_papers/TipSheet_MinimizingBias.pdf)  
