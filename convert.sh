if python3 --version;then
    echo "Python3 Availble"
    if pip check requirements.txt;then
        echo "Requirements satisfied"
        python3 convert.py
    else
        echo "ERROR:Requirements Unsatisfied,quiting"
        echo "Still insist to run?('Y' for yes):"
        read -n1 input 
        if [ "${input}" == "Y" ];then
            python3 convert.py
        else
            echo "Quitting"
        fi
    fi
else
    echo "No Python3 Available,quitting"
    exit 1
fi
