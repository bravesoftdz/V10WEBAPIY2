unit QMenuOBJ;

interface

uses {$IFDEF EAGLCLIENT} Maineagl, {$ELSE EAGLCLIENT} FE_Main, {$ENDIF}
     BPBasic, BPUtil, InitialisationPrev, UFormuleMulBP,
     QUFGBPHPR_TOF, BPPrevCmd, QUFCBPBUDGETPREV_TOF, QUFGBPGRAPHHPR_TOF,
     QUFGBPLOI_TOF, QUFMBPMUL_TOF, QUFSBPARBREPREV_TOF, QUFSBPBUDGETR3_TOF,
     QUFSLARBRE_TOF, QUFVBPARBREMODIF_TOF, QUFVBPDUPLIQUER_TOF, QBPARBRE_TOM,
     QBPLOI_TOM, QUFQBPBUDGET_TOF, QBPCALENDREP_TOM, QUFVBPCALENDJOURM_TOF,
     QUFVBPDETAILJOURM_TOF, QUFVBPDETCALEND_TOF;

Procedure ObjMenuDispatch (Num : Integer);
Procedure ObjModeMenuDispatch (Num : Integer);

implementation

Procedure ObjMenuDispatch (Num : Integer);
begin
  Case Num of
  { 216100 : Objectifs }
    216110 : AglLanceFiche('Q', 'QUFMBPSESSIONENT', '', '', ''); { Session }
    216120 : AglLanceFiche('Q', 'QUFCBPBUDGETPREV', '', '', '');  { Cube }
    216130 : AglLanceFiche('Q', 'QUFSBPBUDGETR3', '', '', '');  { Tableau de bord }
    216141 : AglLanceFiche('Q', 'QUFQBPBUDGET', '', '', 'CONTEXTE=PERIODE'); { Editions }
    { 216200 : Prévisions }
  end;
end;

Procedure ObjModeMenuDispatch (Num : Integer);
begin
  Case Num of
  { 216100 : Objectifs }
    216110 : AglLanceFiche('Q', 'QUFMBPSESSIONENT', '', '', '');   { Session }
    216120 : AglLanceFiche('Q', 'QUFCBPBUDGETPREV', '', '', '');   { Cube }
    216130 : AglLanceFiche('Q', 'QUFSBPBUDGETR3', '', '', '');   { Tableau de bord }
    216150 : AglLanceFiche('Q', 'QUFMBPCALENDREP','','','');   { Calendrier }
    216160 : AglLanceFiche('Q', 'QUFQBPBUDGET', '', '', 'CONTEXTE=PERIODE');  {  Restitution }
    216170 : AglLanceFiche('Q', 'QUFSBPOBJBO','','','');  {  Restitution }
  end;
end;

end.

