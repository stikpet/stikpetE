# stikpetE
The goal of stikpetE is to provide functions for statistical analysis of surveys. 

The functions are also available as a Python package (stikpetP) or r library (stikpetR). For documentation of the functions, you can look at the online documentation of the Python library, or a pdf from the R library. These are available at https://peterstatistics.com/CrashCourse/functions.html.

The only adjustment then should be that a 'dataframe' should be a range in Excel. The Excel add-in does not have functions for visualisations. These can be made with Excel itself usually quite easily.

The functions are NOT optimized but are instead relatively easy to follow. If you would look at the source code it is most often a combination of if statements and for loops.

Of course I could have made mistakes, so use at own risk. 

## Installation

1. Download the stikpetE.xlam file
2. Open Excel
3. Go to File -> Options -> Add-ins -> Go...
4. Click on Browse... and go to the downloaded file

The following functions are included:

|Function|Performs|
|--------|--------|
|di_bsncdf|Bivariate Normal Distribution|
|di_kendall_tau|Kendall Tau Distribution|
|di_mpmf|Multinomial Distribution|
|di_mwwcdf|Mann-Whitney-Wilcoxon Distribution|
|di_qcdf|Studentized Range Distribution|
|di_scdf|Spearman Rho Distribution|
|di_wcdf|Wilcoxon Distribution|
|es_alt_ratio|Alternative Ratio|
|es_bag_s|Bennett-Alpert-Goldstein S|
|es_bin_bin|Effect Sizes for Binary vs. Binary incl. 85 measures|
|es_cohen_d|Cohen d for one-way ANOVA|
|es_cohen_d_os|Cohen d' (for one-sample)|
|es_cohen_d_ps|Cohen d_z (for paired-samples)|
|es_cohen_f|Cohen f|
|es_cohen_g|Cohen g|
|es_cohen_h_os|Cohen h'|
|es_cohen_kappa|Cohen kappa|
|es_cohen_w|Cohen w|
|es_common_language_is|Common Language Effect Size (ind. Samples)|
|es_common_language_ps|Common Language Effect Size (Paired Samples)|
|es_cont_coeff|Contingency Coefficient|
|es_convert|Convert one effect size to another|
|es_cramer_v_gof|Cramér V (for goodness-of-fit)|
|es_cramer_v_ind|Cramér V (for independence)|
|es_dominance|Dominance|
|es_epsilon_sq|Epsilon Squared|
|es_eta_sq|Eta Squared|
|es_freeman_theta|Freeman Theta|
|es_glass_delta|Glass Delta|
|es_goodman_kruskal_lambda|Goodman-Kruskal Lambda|
|es_goodman_kruskal_tau|Goodman-Kruskal tau|
|es_hedges_g_is|Hedges g (ind. samples)|
|es_hedges_g_os|Hedges g (one-sample)|
|es_hedges_g_ps|Hedges g (paired samples)|
|es_jbm_e|Johnston-Berry-Mielke E|
|es_jbm_r|Berry-Johnston-Mielke R|
|es_kendall_w|Kendall w|
|es_odds_ratio|Odds Ratio|
|es_omega_sq|Omega Square|
|es_pairwise_bin|pairwise binary effect sizes|
|es_rmsse|Root Mean Square Standardized Effect Size|
|es_scott_pi|Scott pi|
|es_theil_u|Theil U|
|es_tschuprow_t|Tschuprow T|
|es_vargha_delaney_a|Vargha and Delaney A|
|me_consensus|Consensus|
|me_mean|Mean (incl. different types of means)|
|me_median|Median|
|me_mode|Mode|
|me_mode_bin|Mode for binned data|
|me_quantiles|Quantiles (incl. 18 different methods)|
|me_quartile_range|Quartile Ranges|
|me_quartiles|Quartiles (incl. 20 different methods)|
|me_qv|Qualitative Variation (25 different measures)|
|me_variation|Quantitative Variation (8 different measures)|
|ph_binomial|Pairwise Binomial test|
|ph_column_proportion|Column Proportion test|
|ph_conover_iman|Post-Hoc Conover-Iman Test|
|ph_dunn|Post-Hoc Dunn test (after Kruskal-Wallis)|
|ph_dunn_q|Post-Hoc Dunn test (after Cochran Q)|
|ph_friedman|Post-Hoc Friedman test|
|ph_mcnemar_co|Pairwise McNemar test (collapsing others)|
|ph_mcnemar_pw|Pairwise McNemar test|
|ph_nemenyi|Post-Hoc Nemenyi Test|
|ph_pairwise_is|Post-Hoc Pairwise Independent Samples Test|
|ph_pairwise_iso|Post-Hoc Pairwise Independent Samples Test for Ordinal data|
|ph_pairwise_ps|Post-Hoc Pairwise Paired Samples Tests|
|ph_pairwise_t|Post-Hoc Pairwise Student T|
|ph_residual|Post-Hoc Residual Test|
|ph_sdcf|Post-Hoc Steel-Dwass-Critchlow-Fligner Test|
|r_goodman_kruskal_gamma|Goodman-Kruskal Gamma|
|r_kendall_tau|Kendall Tau (a and b)|
|r_pearson|Pearson Correlation Coefficient|
|r_point_biserial|Point Biserial Correlation Coefficient|
|r_rank_biserial_is|(Glass) Rank Biserial Correlation / Cliff Delta|
|r_rank_biserial_os|Rank biserial correlation coefficient (one-sample)|
|r_rosenthal|Rosenthal Correlation Coefficient|
|r_somers_d|Somers’ d|
|r_spearman_rho|Spearman Rho / Rank Correlation Coefficient|
|r_stuart_tau|Stuart Tau c / Kendall Tau c|
|r_tetrachoric|Tetrachoric Correlation Coefficient|
|tab_cross|Cross Table / Contingency Table|
|tab_frequency|Frequency Table|
|tab_frequency_bins|Binned Frequency Table|
|tab_mult_resp|Multiple Response Table|
|tab_nbins|Number of Bins|
|th_cohen_d|Rules of Thumb for Cohen d|
|th_cohen_g|Rule-of-Thumb for Cohen g|
|th_cohen_h|Rule-of-Thumb for Cohen h|
|th_cohen_w|Rule-of-Thumb for Cohen w|
|th_odds_ratio|Rules of thumb for Odds Ratio|
|th_pearson_r|Rules of Thumb for Pearson Correlation Coefficient|
|ts_alexander_govern_owa|Alexander-Govern Test|
|ts_bhapkar|Bhapkar Test|
|ts_binomial_os|One-Sample Binomial Test|
|ts_box_owa|Box F-Test|
|ts_brown_forsythe_owa|Brown-Forsythe Means Test|
|ts_cochran_owa|Cochran One-Way ANOVA|
|ts_cochran_q|Cochran Q Test|
|ts_fisher|Fisher Exact test|
|ts_fisher_owa|Fisher/Classic One-Way ANOVA / F-Test|
|ts_fligner_policello|Fligner-Policello Test|
|ts_freeman_tukey_gof|Freeman-Tukey Test of Goodness-of-Fit|
|ts_freeman_tukey_ind|Freeman-Tukey Test of Independence|
|ts_freeman_tukey_read|Freeman-Tukey-Read Test of Goodness-of-Fit|
|ts_friedman|Friedman Test|
|ts_g_gof|G (Likelihood Ratio) Test of Goodness-of-Fit|
|ts_g_ind|G (Likelihood Ratio / Wilks) Test of Independence|
|ts_ham_owa|Hartung-Argaç-Makambi Test|
|ts_james_owa|James One-Way Test|
|ts_kruskal_wallis|Kruskal-Wallis H Test|
|ts_mann_whitney|Mann-Whitney U Test|
|ts_mcnemar_bowker|(McNemar-)Bowker Test|
|ts_mehrotra_owa|Mehrotra Test|
|ts_mod_log_likelihood_gof|Mod-Log Likelihood Test of Goodness-of-Fit|
|ts_mod_log_likelihood_ind|Mod-Log Likelihood Test of Independence|
|ts_mood_median|Mood Median Test|
|ts_multinomial_gof|Exact Multinomial Test of Goodness-of-Fit|
|ts_neyman_gof|Neyman Test of Goodness-of-Fit|
|ts_neyman_ind|Neyman Test of Independence|
|ts_ozdemir_kurt_owa|Özdemir-Kurt Test|
|ts_pearson_gof|Pearson Chi-Square Test of Goodness-of-Fit|
|ts_pearson_ind|Pearson Chi-Square Test of Independence|
|ts_powerdivergence_gof|Power Divergence Goodness-of Fit Tests|
|ts_powerdivergence_ind|Power Divergence Test of Independence|
|ts_score_os|One-Sample Score Test|
|ts_scott_smith_owa|Scott-Smith Test|
|ts_sign_os|one-sample sign test|
|ts_sign_ps|Paired Samples Sign Test|
|ts_stuart_maxwell|Stuart-Maxwell / Marginal Homogeneity Test|
|ts_student_t_is|Student t Test (Independent Samples)|
|ts_student_t_os|One-Sample Student t-Test|
|ts_student_t_ps|Student t Test (Paired Samples)|
|ts_trimmed_mean_is|Independent Samples Trimmed/Yuen Mean Test|
|ts_trimmed_mean_os|One-Sample (Yuen or Yuen-Welch) Trimmed Mean Test|
|ts_trinomial_os|One-Sample Trinomial Test|
|ts_trinomial_ps|Trinomial Test (Paired Samples)|
|ts_wald_os|One-Sample Wald Test|
|ts_welch_owa|Welch One-Way ANOVA|
|ts_welch_t_is|Welch t Test (Independent Samples)|
|ts_wilcox_owa|Wilcox Test|
|ts_wilcoxon_os|One-Sample Wilcoxon Signed Rank Test|
|ts_wilcoxon_ps|Paired Samples Wilcoxon Signed Rank Test|
|ts_z_is|Independent Samples Z Test|
|ts_z_os|One-Sample Z Test|
|ts_z_ps|Z-test (Paired Samples)|

