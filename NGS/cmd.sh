#!/usr/bin/bash

bcftools mpileup -Ou -f Homo_sapiens.GRCh38.dna.primary_assembly.fa.gz -a AD,DP -R genes_region.txt -b bam.list.txt | bcftools call --threads 8 -mv -a PV4,GQ -Ob -o cohort_231.bcf
