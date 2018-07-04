{***********UNITE*************************************************
Auteur  ...... : BM
Créé le ...... : 05/11/2002
Modifié le ... :   /  /
Description .. : Source TOM Générique et TOM héritantes pour les
                 TABLES : TOM_JUFORMEJUR, TOM_JUTYPECIVIL,
                 TOM_JUTYPEEVT, TOM_JUTYPEPER
Mots clefs ... : TOM;JUFORMEJUR;JUTYPECIVIL;JUTYPEEVT;JUTYPEPER
*****************************************************************}

Unit UTOMTypeCivil;
//////////////////////////////////////////////////////////////////
Interface
//////////////////////////////////////////////////////////////////
Uses
{$IFNDEF EAGLCLIENT}
{$ENDIF}
     Classes, UTOMTypeParam;
//////////////////////////////////////////////////////////////////
Type
   TOM_JUTYPECIVIL = Class (TOM_TYPEPARAM)
      private
         procedure FILTRESEXES_OnClick( Sender : TObject );
         procedure FiltreInitialiser ( sTypeSexe_p : string ) ;
      public
         procedure OnArgument ( sParams_p : String )   ; override;
   end ;
//////////////////////////////////////////////////////////////////
Implementation
//////////////////////////////////////////////////////////////////
uses
   hctrls;
//////////////////////////////////////////////////////////////////
{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. MERIAUX
Créé le ...... : 22/07/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOM_JUTYPECIVIL.OnArgument ( sParams_p : String ) ;
begin
   Inherited ;
   sTypeElement_f := 'de civivilité';
   THRadioGroup(GetControl('FILTRESEXES')).OnClick := FILTRESEXES_OnClick;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B.MERIAUX
Créé le ...... : 22/10/2003
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOM_JUTYPECIVIL.FILTRESEXES_OnClick( Sender : TObject );
var
   sCodeSexe_l : string;
begin
   sCodeSexe_l := THRadioGroup(GetControl('FILTRESEXES')).Value;
   FiltreInitialiser( sCodeSexe_l );
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B.MERIAUX
Créé le ...... : 22/10/2003
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOM_JUTYPECIVIL.FiltreInitialiser ( sTypeSexe_p : string ) ;
var
   sFiltreListe_l : string;
begin
     // $$$ JP 26/11/2004 - dans la tablette YYSEXE, il n'y a que: H, F et I (pas M)
   if sTypeSexe_p = '' then   // Tous
        sFiltreListe_l := ''
   else if sTypeSexe_p = 'I' then   // Indifférenciés
        sFiltreListe_l := 'JTC_SEXE = '''''
   else if sTypeSexe_p = 'M' then   // Hommes
        sFiltreListe_l := 'JTC_SEXE = ''M''' // OR JTC_SEXE = ''M'')'
   else if sTypeSexe_p = 'F' then   // Femmes
        sFiltreListe_l := 'JTC_SEXE = ''F''';

   FiltreActiver(sFiltreListe_l);
end;

//////////////////////////////////////////////////////////////////

Initialization
   registerclasses ( [ TOM_JUTYPECIVIL ] ) ;
end.
