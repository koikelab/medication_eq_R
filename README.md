# medication_eq_R

An R code for calculating equivalent doses especially for psychiatry research [1]. An example sourse list is made for chlorpromazine, imipramine, diazepam, and biperiden equivalent doses in Japanese medical situation [2,3]. Therefore, please refer to this article when using your research. The code can apply to different equivalent doses and drug names suitable for each research site. 

## References

[1] Koike S, Hirano Y, Nakajima S, Morita K. Medication equivalent dose program in R. Psychiatry and Clinical Neurosciences. 2025 in press.<br>
[2] Inada T, Inagaki A. Psychotropic dose equivalence in Japan. Psychiatry and Clinical Neurosciences. 2015; 69: 440–7. 
[3] Inagaki A, Inada T. Dose equivalence of psychotropic drugs (Part 28). Dose equivalence of novel antipsychotics : Brexpiprazole. Japanese Journal of Clinical Psychopharmacology. 2022; 25: 91–8. 

## How to use
1. Download and extract all files and folders.
2. Lanunch R program in 'R' folder.
3. In R studio, run 'source'.
4. When running the program, you can choose an input file. For example, you choose '../input/example.xlsx'.
5. The calculated equivalent doses per row and per ID and date in the 'output' folder.
6. If you want to change the drug name and equivalent dose pair in the source file, you choose the edited file in Line 10 and 14 of the R code. If you want to add the new equivalent dose, please refer to Line 29-38 in the code.

## Update history

### Oct 4, 2025
A minor revision of medication_list.xlsx in the source folder.<br>
Edit the R code for the source file.

