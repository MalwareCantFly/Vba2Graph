rm -rf olevba_output
rm -rf output

mkdir olevba_output

for file in input/*; do
	echo "olevba $file"
   	olevba3 -c "$file" > "olevba_output/$(basename $file).bas"
done

for file in olevba_output/*; do
    python3 vba2graph.py -i "$file" -o "output"
done
