---
title: Tests d?utilisation pour les compl?ments Office
description: ''
ms.date: 01/23/2018
ms.openlocfilehash: 410b8d7ede22cf222ee2df794e438c7f5f8881dd
ms.sourcegitcommit: c72c35e8389c47a795afbac1b2bcf98c8e216d82
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/23/2018
---
# <a name="usability-testing-for-office-add-ins"></a>Tests d?utilisation pour les compl?ments Office

Une excellente conception de compl?ment tient compte des comportements des utilisateurs. ?tant donn? que vos propres id?es pr?con?ues influenceront vos d?cisions de conception, il est important de tester les conceptions avec des utilisateurs r?els pour vous assurer que vos compl?ments fonctionnent correctement pour vos clients. 

Vous pouvez ex?cuter les tests d?utilisation de diff?rentes fa?ons. Pour de nombreux d?veloppeurs de compl?ments, les ?tudes d?utilisation ? distance sans mod?rateur sont les plus rentables et les plus rapides. Plusieurs services de test connus facilitent cette t?che ; en voici quelques exemples : 

 - [UserTesting.com](https://www.UserTesting.com)
 - [Optimalworkshop.com](https://www.Optimalworkshop.com)
 - [Userzoom.com](https://www.Userzoom.com)

Ces services de test vous aident ? simplifier la cr?ation d?un plan de test et ?liminent le besoin de rechercher des participants ou de mod?rer les tests. 

Vous avez seulement besoin de cinq participants pour r?v?ler la plupart des probl?mes d?utilisation dans votre conception. Incorporez r?guli?rement des petits tests dans votre cycle de d?veloppement pour vous assurer que votre produit est centr? sur l?utilisateur.

> [!NOTE]
> Nous vous recommandons de tester l?utilisation de votre compl?ment sur plusieurs plateformes. Pour [publier votre compl?ment dans AppSource](https://docs.microsoft.com/en-us/office/dev/store/submit-to-the-office-store), il doit fonctionner sur toutes les [plateformes qui prennent en charge les m?thodes que vous d?finissez](../overview/office-add-in-availability.md).

## <a name="1---sign-up-for-a-testing-service"></a>1.   Inscrivez-vous ? un service de test

Pour plus d?informations, reportez-vous ? la section sur la [s?lection d?un outil en ligne pour les tests utilisateur distants sans mod?rateur](https://www.nngroup.com/articles/unmoderated-user-testing-tools/).

## <a name="2-develop-your-research-questions"></a>2. D?veloppez vos questions de recherche
 
Les questions de recherche d?finissent les objectifs de votre recherche et guident votre plan de test. Vos questions vous aideront ? identifier les participants ? recruter et les t?ches qu?ils ex?cuteront. R?digez vos questions de recherche avec autant de pr?cision que possible. Vous pouvez ?galement rechercher des r?ponses ? des questions plus larges.
 
Voici quelques exemples de questions de recherche :
  
**Sp?cifiques**  

 - Les utilisateurs remarquent-ils le lien indiquant ? version d??valuation gratuite ? sur la page d?accueil ?
 - Lorsque les utilisateurs ins?rent du contenu dans leur document ? partir du compl?ment, savent-ils o? il est ins?r? dans le document ?

**Larges**  

 - Quelles sont les difficult?s majeurs pour l?utilisateur dans notre compl?ment ?
 - Les utilisateurs comprennent-ils la signification des ic?nes dans notre barre de commandes avant de cliquer dessus ?
 - Le menu des param?tres est-il facilement accessible pour les utilisateurs ?

Il est important d?obtenir des donn?es sur l?int?gralit? du parcours des utilisateurs, de la d?couverte de votre compl?ment jusqu?? son installation et son utilisation. Envisagez de r?diger des questions de recherche qui abordent les aspects suivants de l?exp?rience utilisateur dans le compl?ment :
 
 - Recherche de votre compl?ment dans AppSource
 - D?cision d?installation de votre compl?ment
 - Premi?re ex?cution
 - Commandes du ruban
 - Interface utilisateur du compl?ment
 - Interaction du compl?ment avec l?espace d?di? aux documents de l?application Office
 - Niveau de contr?le de l?utilisateur sur les flux d?insertion de contenu

Pour plus d?informations, voir la section sur la [r?daction de questions efficaces](http://help.usertesting.com/customer/en/portal/articles/2077663-writing-effective-questions).
 
## <a name="3-identify-participants-to-target"></a>3. Identifiez les participants ? cibler
 
Les services de test ? distance peuvent permettre de contr?ler de nombreuses caract?ristiques de vos participants aux tests. R?fl?chissez soigneusement aux types d?utilisateurs que vous souhaitez cibler. Lors des premi?res ?tapes de collecte de donn?es, il peut ?tre pr?f?rable de recruter un large ?ventail de participants diff?rents pour identifier les probl?mes d?utilisation plus ?vidents. Plus tard, vous pourrez d?cider de cibler des groupes comme les utilisateurs Office avanc?s, des professions en particulier ou des tranches d??ge sp?cifiques.
 
## <a name="4-create-the-participant-screener"></a>4. Cr?ez le filtre de participants
 
Le filtre est l?ensemble de questions et d?exigences que vous allez pr?senter aux participants potentiels des tests afin de les filtrer. N?oubliez pas que les participants pour les services tels que UserTesting.com sont financi?rement motiv?s pour ?tre s?lectionn?s pour votre test. Il est conseill? d?inclure des questions pi?ge dans votre filtre si vous souhaitez exclure certains utilisateurs des tests. 
 
Par exemple, si vous recherchez des participants qui connaissent GitHub, pour exclure les utilisateurs qui mentent, incluez des r?ponses qui ne vous int?ressent pas dans la liste des r?ponses possibles.

**Parmi les r?f?rentiels de code source suivants, lesquels connaissez-vous ?**  
 a. SourceShelf  [*R?ponse disqualifiante*]  
 b. CodeContainer  [*R?ponse disqualifiante*]  
 c. GitHub  [*Doit s?lectionner cette r?ponse*]  
 d. BitBucket  [*Peut s?lectionner cette r?ponse*]  
 e. CloudForge  [*Peut s?lectionner cette r?ponse*]  

Si vous envisagez de tester une version d?j? en ligne de votre compl?ment, les questions suivantes peuvent permettre de s?lectionner les utilisateurs qui seront en mesure de le faire. 

**Ce test exige que vous disposiez de Microsoft PowerPoint 2016. Avez-vous PowerPoint 2016 ?**  
 a. Oui [*Doit s?lectionner cette r?ponse*]  
 b. Non [*R?ponse disqualifiante*]  
 c. Je ne sais pas [*R?ponse disqualifiante*]  

**Pour ce test, vous devez installer un compl?ment gratuit pour PowerPoint 2016 et cr?er un compte gratuit pour l?utiliser. ?tes-vous pr?t ? installer un compl?ment et ? cr?er un compte gratuit ?**  
 a. Oui [*Doit s?lectionner cette r?ponse*]  
 b. Non [*R?ponse disqualifiante*]  

Pour plus d?informations, consultez les [meilleures pratiques en mati?re de questions de filtrage](http://help.usertesting.com/customer/en/portal/articles/2077835-screener-question-best-practices).
 
## <a name="5-create-tasks-and-questions-for-participants"></a>5. Cr?ez des t?ches et des questions pour les participants
 
Essayez de classer ce que vous voulez tester par ordre de priorit? afin de limiter le nombre de t?ches et de questions pour le participant. Certains services paient les participants uniquement pour une certaine dur?e ; veillez donc ? ne pas d?passer ce temps.

Autant que possible, essayez d?observer les comportements des participants au lieu de les interroger sur leurs comportements. Si vous avez besoin les interroger sur leurs comportements, demandez-leur ce qu?ils ont d?j? fait, plut?t que ce qu?ils feraient dans telle ou telle situation. De cette fa?on, les r?sultats seront plus fiables.
 
La principale difficult? lors des tests sans mod?rateur consiste ? s?assurer que vos participants comprennent vos t?ches et vos sc?narios. Vos instructions doivent ?tre *claires et concises*. In?vitablement, si vos instructions ne sont pas claires, certains participants ne les comprendront pas. 

Ne supposez jamais que l?utilisateur sera sur l??cran o? il est cens? ?tre pendant le test. Vous pouvez lui indiquer l??cran sur lequel il doit se trouver au d?but de la t?che suivante. 

Pour plus d?informations, voir la section expliquant [comment r?diger des instructions efficaces pour les t?ches](http://help.usertesting.com/customer/en/portal/articles/2077824-writing-great-tasks).

## <a name="6-create-a-prototype-to-match-the-tasks-and-questions"></a>6. Cr?ez un prototype pour faire correspondre les t?ches et les questions
 
Vous pouvez faire tester votre compl?ment d?j? en ligne ou vous pouvez faire tester un prototype. N?oubliez pas que si vous souhaitez tester un compl?ment d?j? en ligne, vous devez filtrer les participants pour ne s?lectionner que ceux qui ont Office 2016, qui sont pr?ts ? installer le compl?ment et qui sont pr?ts ? cr?er un compte (sauf si vous leur fournissez des informations d?identification). Vous devez ensuite pour vous assurer qu?ils installent correctement votre compl?ment. 

En moyenne, aider un utilisateur ? installer un compl?ment prend environ 5 minutes. Voici un exemple d??tapes d?installation claires et concises. Modifiez-les en fonction des caract?ristiques propres ? votre test.

**Installez le compl?ment (indiquez le nom de votre nom compl?ment ici) pour PowerPoint 2016, en suivant les instructions suivantes :** 

1. Ouvrez Microsoft PowerPoint 2016.
2. S?lectionnez **Nouvelle pr?sentation**.
3. Acc?dez ? **Ins?rer > Mes compl?ments**.
5. Dans la fen?tre contextuelle, choisissez **Magasin**.
6. Saisissez (nom du compl?ment) dans la zone de recherche.
7. Choisissez (nom du compl?ment).
8. Prenez quelques instants pour examiner la page du magasin et pour vous familiariser avec le compl?ment.
9. Choisissez **Ajouter** pour installer le compl?ment.

Vous pouvez tester un prototype ? n?importe quel niveau d?interaction et de fid?lit? visuelle. Pour une liaison et une interactivit? plus complexes, pensez ? utiliser un outil de prototypage tel que [InVision](https://www.invisionapp.com). Si vous souhaitez simplement tester des ?crans statiques, vous pouvez h?berger les images en ligne et envoyer l?URL correspondante aux participants, ou leur donner un lien vers une pr?sentation PowerPoint en ligne. 

## <a name="7-run-a-pilot-test"></a>7. Effectuez un test pilote

Il peut ?tre difficile de mettre au point le prototype appropri? et la liste de t?ches/question ad?quate. Les utilisateurs peuvent ne pas comprendre certaines t?ches, ou se perdre dans votre prototype. Vous devez ex?cuter un test pilote avec 1 ? 3 utilisateurs pour solutionner les probl?mes in?vitables au niveau du format du test. Cette op?ration permet de s?assurer que vos questions sont claires, que le prototype est correctement configur? et que vous allez pouvoir recueillir le type de donn?es que vous recherchez.

## <a name="8-run-the-test"></a>8. Lancez le test

Une fois que vous avez command? votre test, vous obtenez des notifications par courrier ?lectronique lorsque les participants l?effectuent. Sauf si vous avez cibl? un groupe sp?cifique de participants, les tests sont g?n?ralement effectu?s en quelques heures.

## <a name="9-analyze-results"></a>9. Analysez les r?sultats

Vous devez maintenant essayer d?interpr?ter les donn?es que vous avez collect?es. Pendant que vous regardez les vid?os des tests, notez les probl?mes que rencontre l?utilisateur, ainsi que les points positifs. N?essayez pas d?interpr?ter la signification des donn?es tant que vous n?avez pas affich? tous les r?sultats. 

Un probl?me d?utilisation rencontr? par un seul participant n?est pas suffisant pour justifier une modification de la conception. Deux ou plusieurs participants rencontrant le m?me probl?me sugg?re que d?autres utilisateurs dans la population globale rencontreront ?galement ce probl?me.

En r?gle g?n?rale, soyez prudent lorsque vous utilisez vos donn?es pour tirer des conclusions. N?essayez pas d?interpr?ter les donn?es de sorte qu?elles aillent dans un sens en particulier. Ne tombez pas dans ce pi?ge. Soyez honn?te lorsque vous identifiez ce que les donn?es prouvent r?ellement ou ne prouvent pas, et n?h?sitez pas ? reconna?tre que, parfois, elles ne procurent aucune information exploitable. Gardez l?esprit ouvert. Les comportements des utilisateurs vont souvent ? l?encontre des attentes du concepteur.
 

## <a name="see-also"></a>Voir aussi
 
 - [R?alisation de tests d?utilisation](http://whatpixel.com/howto-conduct-usability-testing/)  
 - [Meilleures pratiques](http://help.usertesting.com/customer/en/portal/articles/1680726-best-practices)  
 - [R?duction de la subjectivit?](http://downloads.usertesting.com/white_papers/TipSheet_MinimizingBias.pdf)  
