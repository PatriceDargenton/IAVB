Intelligence Artificielle Vraiment Basique (IAVB)

# Version 3.12 du 18/06/2017
- IAVB : Intelligence Artificielle en Visual Basic -> Vraiment Basique (pour tenir compte des versions en C#) ;
- Synth�se vocale en fran�ais de Windows 10 (sous Windows 7 on n'a que l'anglais), en plus de MS-Agent ;
- Version VB.Net : mode automatique pour tester automatiquement tous les exemples, zone de liste pour la synth�se vocale ;
- Gestion des majuscules : ajout d'un bool�en (bTraiterEnMinuscules) pour pouvoir comparer les r�sultats avec la version d'origine (si on active les accents, par exemple pour que la synth�se vocale fonctionne bien, alors il faut d�sactiver les majuscules, car les accents ne sont pas faciles � trouver au clavier sur les majuscules) ;
- Gestion des accents : ajout d'un bool�en (bTraiterSansAccents) pour pouvoir comparer les r�sultats avec la version d'origine ;
- Normalisation des r�ponses : ajout d'un bool�en (bNormalisationSortieTrimEtVbLf) pour rendre homog�ne les espaces et sauts de ligne dans les r�ponses, on avait parfois saut de ligne en trop, ou espace en trop, parfois il en manquait ;
- Version en VBA Word (code commun avec la version VB6), car il est devenu difficile de coder en VB6, l'IDE de VB6 ne s'installant pas sous Windows 10 ;
- Version VBA d'origine : fonctionnement identique � la version de 1984 (reprise des exemples pour pouvoir v�rifier qu'on obtient bien les m�mes r�sultats) ;
- Versions VBA : correction de la copie dans le presse-papier sous Windows 64 bit (l'ancienne m�thode ne fonctionnait que sous Windows 32 bits) ;
- Versions en C# : WinForm, WPF et CSharpHtml (tout juste fonctionnel, mais � terminer, il manque notamment ListBox.ScrollIntoView) ;
- Tests unitaires pour v�rifier les exemples fournis ;
- Respect des r�gles d'analyse du code selon Visual Studio 2017 ;
- Mise au norme DotNet du code h�rit� de VB6 (utilisation de Generic.List(Of String), du coup on part de l'index 0 et non plus 1, ...) ;
- Version sans Microsoft.VisualBasic.dll (pour pouvoir convertir plus facilement en C#) ;
- Passage en VB 2013 (les versions 2015 et 2017 ont toujours un bug d'indentation, si on a utilis� l'indentation en mode bloc ou bien d�sactiv�).

[IAVB.html](http://patrice.dargenton.free.fr/ia/iavb/IAVB.html) : Documentation compl�te