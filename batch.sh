bash -c "rm olevba_output/*"
bash -c "rm -rf output/*"

for file in input/*; do
	echo "olevba $file"
   	olevba -c "$file" > "olevba_output/$(basename $file).bas"
done

for file in olevba_output/*; do
    python vba2graph.py -i "$file" -o "output"
done