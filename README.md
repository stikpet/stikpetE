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

## Version 0.0.2

This contains all my functions for 'basic' univariate analysis, it has the following functions available:

Correlations

1. r_rank_biserial_os()
1. r_rosenthal()

Effect sizes

1. es_convert()
1. es_alt_ratio()
1. es_cohen_d_os()
1. es_cohen_g()
1. es_cohen_h_os()
1. es_cohen_w()
1. es_cramer_v_gof()
1. es_dominance()
1. es_hedges_g_os()
1. es_jbm_e()

Measures

1. me_concensus()
1. me_mean()
1. me_median()
1. me_mode()
1. me_mode_bin()
1. me_quartile_range()
1. me_quartiles()
1. me_qv()
1. me_variation_ratio()

Other Functions

1. ph_binomial()
1. tab_frequency()
1. tab_frequency_bins()
1. tab_nbins()
1. th_cohen_d()
1. th_cohen_g()
1. th_cohen_h()
1. th_cohen_w()
1. th_pearson_r()

Tests

1. ts_binomial_os()
1. ts_freeman_tukey_gof
1. ts_freeman_tukey_read()
1. ts_g_gof()
1. ts_mod_log_likelihood_gof()
1. ts_multinomial_gof()
1. ts_neyman_gof()
1. ts_pearson_gof()
1. ts_powerdivergence_gof()
1. ts_score_os()
1. ts_sign_os()
1. ts_student_t_os()
1. ts_trimmed_mean_os()
1. ts_trinomial_os()
1. ts_wald_os()
1. ts_wilcoxon_os()
1. ts_z_os()

