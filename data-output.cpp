#include <iostream>
#include <fstream>
#include <map>
#include <set>
#include <vector>
#include <string>
#include <cstdlib>
using namespace std;
// 读取班级列表并存储在容器中
vector<string> readClassList(const string& classListFileName) {
    vector<string> classList;
    ifstream file(classListFileName);
    if (file.is_open()) {
        string className;
        while (getline(file, className)) {
            classList.push_back(className);
        }
        file.close();
    }
    return classList;

}
//写入data.txt
void writeDataToFile(const string& fileName, int weekNumber, const map<int, vector<string>>& absentData, const map<int, vector<string>>& lateData,const map<int, vector<string>>& leaveData) {
    ofstream file(fileName);
    if (file.is_open()) {
        file << "周数:" << weekNumber << endl;

        if (!absentData.empty()) {
            file << "旷课" << endl;
            for (const auto& entry : absentData) {
                for (const string& info : entry.second) {
                    file << info << endl;
                }
            }
        }

        if (!lateData.empty()) {
            file << "迟到" << endl;
            for (const auto& entry : lateData) {
                for (const string& info : entry.second) {
                    file << info << endl;
                }
            }
        }
        if (!leaveData.empty()) {
            file << "请假" << endl;
            for (const auto& entry : leaveData) {
                for (const string& info : entry.second) {
                    file << info << endl;
                }
            }
        }
        file.close();
        cout << "数据已写入文件 " << fileName << endl;
    } else {
        cerr << "无法打开文件 " << fileName << " 来写入数据." << endl;
    }
}

int main() 
{
    //system("chcp 65001");
    int weekNumber;
    string currentClass;  // 用于跟踪当前的班级
    map<int, vector<string>> absentData;  // 旷课信息容器
    map<int, vector<string>> lateData;    // 迟到信息容器
    map<int, vector<string>> leaveData;    // 迟到信息容器
    cout << "请输入周数：";
    cin >> weekNumber;
    int cnt = 0;
    string data;
    //int flag=0;
    int cnt_k=0;
    int flag_text=0;
    // 读取班级列表
    vector<string> classList = readClassList("class.txt");
    set<string> recordedClasses;
    string str;
    while (true) 
    { //cnt==0，就是第一次录入此班的数据
        if (cnt == 0) {
            cout << "班级：";
            cin >> currentClass;
            data = currentClass + "班:";
            flag_text=0; 
        }//开始录入班级，进行操作选择
        if(currentClass=="-1")
        {
            data.clear();
            break;
        }
        int choice;
        if(cnt==0)
        {
            cout << "请选择操作 (1: 旷课, 2: 迟到, 3: 请假 ): ";
            cin>>choice;
            cout<<"-----------------------------------------"<<endl;
            if(choice==1)
            {
            	str="旷课";
            	cout<<currentClass<<"班"<<str<<"录入："<<endl;
            }else if(choice==2)
            {
            	str="迟到";
            	cout<<currentClass<<"班"<<str<<"录入："<<endl;
            }else if(choice==3)
            {
                str="请假";
            	cout<<currentClass<<"班"<<str<<"录入："<<endl;
            }
            
        }//如果录完了就再进行选择，选择旷课或迟到
        
        if(choice==1)   str="旷课";
        if(choice==2)   str="迟到";
        if(choice==3)   str="请假";
        
        if (choice == 0) {
        	cout<<">>"<<"您选择了退出系统"<<endl; 
            break;
        }
        string studentName, dayTime;
        cout << "姓名：";
        cin >> studentName;
        if(choice==2)	flag_text=1; 
        if(choice==3)   flag_text=-1;
        if(studentName=="-1")
        {
        	// 记录已录入的班级
            recordedClasses.insert(currentClass);
            cnt=0;
            data.erase(data.end() - 1);
            cout<<">>"<<currentClass<<"班数据已全部录入完毕！"<<endl;
            cout<<"-----------------------------------------"<<endl;
            cnt_k++;
            cout<<"----您已完成"<<cnt_k<<"个班级的录入，很棒棒哦~-----"<<endl<<endl; 
            if(data.size()<=6)
			{
				data.clear();
				continue;
			} 
            if(flag_text==1)	 lateData[weekNumber].push_back(data);//退出且有迟到数据就把数据放进迟到的容器里面
            else if(flag_text==0) absentData[weekNumber].push_back(data);
            else if(flag_text==-1) leaveData[weekNumber].push_back(data);
            data.clear();
            continue;
        }
         if(studentName=="1")
        {
            choice=1;
            data.erase(data.end() - 1);
            cout<<">>"<<currentClass<<"班"<<str<<"数据已全部录入完毕！"<<endl;
            cout<<"-----------------------------------------"<<endl;
            if(data.size()<=6)
			{
				data.clear();
				continue;
			}
            cout<<currentClass<<"班旷课录入："<<endl;
            if(flag_text==1)	 lateData[weekNumber].push_back(data);//退出且有迟到数据就把数据放进迟到的容器里面
            else if(flag_text==0) absentData[weekNumber].push_back(data);
            else if(flag_text==-1) leaveData[weekNumber].push_back(data);
            data.clear();
            flag_text=0;
            data = currentClass + "班:";
            continue;
        }
     
        if(studentName=="2")
        {
            choice=2;//记后续的数据为迟到
            data.erase(data.end() - 1);
            //data.erase(data.end() - 2);
            cout<<">>"<<currentClass<<"班"<<str<<"数据已全部录入完毕！"<<endl;
            cout<<"-----------------------------------------"<<endl;
            cout<<currentClass<<"班迟到录入："<<endl;
            if(data.size()<=6)
			{
				data.clear();
				continue;
			}
            if(flag_text==1)	 lateData[weekNumber].push_back(data);//退出且有迟到数据就把数据放进迟到的容器里面
            else if(flag_text==0) absentData[weekNumber].push_back(data);
            else if(flag_text==-1) leaveData[weekNumber].push_back(data);
            data.clear();
            flag_text=1;
            data = currentClass + "班:";
            continue;
        }
        if(studentName=="3")
        {
            choice=3;
            data.erase(data.end() - 1);
            cout<<">>"<<currentClass<<"班"<<str<<"数据已全部录入完毕！"<<endl;
            cout<<"-----------------------------------------"<<endl;
            if(data.size()<=6)
			{
				data.clear();
				continue;
			}
            cout<<currentClass<<"班请假录入："<<endl;
            if(flag_text==1)	 lateData[weekNumber].push_back(data);//退出且有迟到数据就把数据放进迟到的容器里面
            else if(flag_text==0) absentData[weekNumber].push_back(data);
            else if(flag_text==-1) leaveData[weekNumber].push_back(data);
            data.clear();
            flag_text=-1;
            data = currentClass + "班:";
            continue;
        }
        if(choice==1||choice==2)
        {
            cout << "星期和节次：";
            cin.ignore();
            getline(cin,dayTime);
            data += studentName + ":" + dayTime;
            int len=data.size()-1;
            if(data[len]==' ')		data.erase(data.end() - 1); 
        }else
        {
        	data += studentName;
        }
        data += ",";
        cnt = 1;  // 标志位，延用后就设为1
        //} 
    }


    // 写入数据到文本文件
    writeDataToFile("data.txt", weekNumber, absentData, lateData,leaveData);
	// 输出没有被录入的班级
    cout << "-----------------------------------------" << endl;
    cout << "以下班级未被录入：" << endl;
    int cnt_p=0;
    for (const string& className : classList) {
        if (recordedClasses.find(className) == recordedClasses.end()) {
        	++cnt_p;
            cout << className << "\t";
            if(cnt_p%5==0)	cout<<endl;
        }
    }
    cout<<endl;
    system("pause");
    return 0;
}
    
