
# <a name="labsjs-lab-components"></a>LabsJS lab components

Labs.js vous propose quatre types de composant que vous pouvez utiliser pour assembler votre laboratoire. Chaque type de composant prend en charge un type spécifique d’interaction avec le laboratoire, notamment des problèmes de choix multiples, des problèmes de réponse libre ou des activités comme l’affichage de pages web dans la balise iFrame du code HTML de la leçon.

## <a name="components"></a>Components

Office Mix prend en charge les quatre types de composant de laboratoire suivants : 


-  **Activity component** ( **IActivityComponent**). Presents the user with an activity that must be completed; for example, read a piece of text, watch a video, or interact with a simulation. For more information, see [Labs.Components.ActivityComponentInstance](../../../reference/office-mix/labs.components.activitycomponentinstance.md).
    
-  **Choice component** ( **IChoiceComponent**). Presents the user with a list of choices from which the user must select. Supports single or multiple responses (or no answer at all). Use this component type for true/false, multiple choice, multiple response, or polls. For more information, see [Labs.Components.ChoiceComponentInstance](../../../reference/office-mix/labs.components.choicecomponentinstance.md).
    
-  **Input component** ( **IInputComponent**). Enables free form user input. Use this component type when you want to get responses to questions or math problems from the user, for example, or for other problem types that require text inputs from the user. For more information, see [Labs.Components.InputComponentInstance](../../../reference/office-mix/labs.components.inputcomponentinstance.md).
    
-  **Dynamic component** ( **IDynamicComponent**). Generates other component types at runtime. Use this component type when you have branching questions, for example, where follow-up component types vary depending on a previous user input. This type also enables creating quiz banks or generating problems at runtime. For more information, see [Labs.Components.DynamicComponentInstance](../../../reference/office-mix/labs.components.dynamiccomponentinstance.md).
    

## <a name="additional-resources"></a>Ressources supplémentaires



- [Compléments Office Mix](../../powerpoint/office-mix/office-mix-add-ins.md)
    
- [Configuration et modification de laboratoires LabsJS pour Office Mix](../../powerpoint/office-mix/configuring-and-editing-labsjs-labs-for-office-mix.md)
    
- [Procédure pas à pas : Création de votre premier laboratoire pour Office Mix](../../powerpoint/office-mix/creating-your-first-lab-for-office-mix.md#walkthrough-creating-your-first-lab-for-office-mix)
    
