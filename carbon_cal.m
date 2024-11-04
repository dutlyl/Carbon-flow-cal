
%% ---------------------------------in---------------------------------%%
[line_data]          =    xlsread('Line_data_1',1);
[loss_data]          =   xlsread('Line_data_1',2);
[node_data]        =   xlsread('Node_data_2',1);
[region_data]      =   xlsread('Node_data_2',2);
[load_data]         =   xlsread('Node_data_2',3);
[gen_data]          =   xlsread('Node_data_2',4);
%% -----------------------------------------------------------------%%
for i = 1:24
    carbon(i).line(:,1)    =   line_data(:,3+i);
    carbon(i).load(:,1)   =   load_data(:,3+i);
    carbon(i).gen(:,1)    =   gen_data(:,6+i);
    carbon(i).loss(:,1)    =   loss_data(:,3+i);
end
%% ------------------------------------------------------------------%%
for i = 1:24    %24小时
 %-------------------------------------------------------------------------------------------------------------------------------------------%
    K   =   size(gen_data,1);                 
    N   =   size(node_data,1);                  
    M   =   size(load_data,1);                     
    P     =   size(line_data,1);                   
    Q    =   size(region_data,1);                 
    L     =   zeros(M,N);                              
    J      =   zeros(K,N);                                
    P_B =   ones(N,N)*0.1;                      
    E_N     =   zeros(N,1);                             
    P_N     =   zeros(N,N);                         
    E_G     =   zeros(K,1);                           
    E_G     =   gen_data(:,6);
%------------------------------------------------------------------------------------------------------------------% 
    for j=1:P
        if  carbon(i).line(j,1)>0
            P_B(find(node_data(:,1)==line_data(j,2)),find(node_data(:,1)==line_data(j,3)))    =...
                P_B(find(node_data(:,1)==line_data(j,2)),find(node_data(:,1)==line_data(j,3)))    +   carbon(i).line(j,1);
        else
            P_B(find(node_data(:,1)==line_data(j,3)),find(node_data(:,1)==line_data(j,2)))    =...
                P_B(find(node_data(:,1)==line_data(j,3)),find(node_data(:,1)==line_data(j,2)))    -    carbon(i).line(j,1);
        end
    end
%------------------------------------------------------------------------------------------------------------------%    
     for j = 1:M
        L(j,find(node_data(:,1)==load_data(j,1)))    =   1;
    end
    P_L     =   diag(carbon(i).load(:,1))*L;   
%------------------------------------------------------------------------------------------------------------------%
    for j=1:K
        J(j,find(node_data(:,1)==gen_data(j,2)))    =   1;
    end
    P_G     =   diag(carbon(i).gen(:,1))*J;
%------------------------------------------------------------------------------------------------------------------%
    P_Z     =   [P_B;P_G];                                    
    P_N     =   diag(ones(1,N+K)*P_Z);          
    E_N     =   inv(P_N-P_B')*P_G'*E_G;           
    R_B     =   diag(E_N)*P_B;                        
%------------------------------------------------------------------------------------------------------------------%
    carbon(i).gen_node=sum(P_G,1)';          
    carbon(i).node_E = E_N;                        
    carbon(i).node_CR = E_G' * P_G;                
    carbon(i).node_CU = E_N' * P_L;                
    carbon(i).node_CF = E_N' * P_N;               
    for j=1:P                                                         
        if  carbon(i).line(j,1)>0
            carbon(i).line_E(j,1) = E_N(find(node_data(:,1)==line_data(j,2)));
        else
            carbon(i).line_E(j,1) = E_N(find(node_data(:,1)==line_data(j,3)));
        end
    end
    carbon(i).line_Closs = carbon(i).loss(:,1) .* carbon(i).line_E(:,1);           
    carbon(i).line_CF= carbon(i).line(:,1) .* carbon(i).line_E(:,1);                   
    carbon(i).gen_Cpai = carbon(i).gen(:,1) .* E_G;                                            
     for j= 1:Q
        node_in_region = find(node_data(:,4)==j);   
        gen_node = find(gen_data(:,4)==j);
        carbon(i).area_load(j,1) = sum(carbon(i).load(node_in_region,1));        
        carbon(i).area_gen(j,1)  = sum(carbon(i).gen(gen_node,1));                  
        carbon(i).area_E(j,1) = sum(carbon(i).load(node_in_region,1).*carbon(i).node_E(node_in_region,1))/carbon(i).area_load(j,1);  
        carbon(i).area_CR(j,1) = sum(carbon(i).node_CR(1,node_in_region));    
        carbon(i).area_CU(j,1) = sum(carbon(i).node_CU(1,node_in_region));    
        carbon(i).area_CF(j,1) = sum(carbon(i).node_CF(1,node_in_region));     
     end   
end
%% -------------------------------------------------------------%%

%------------------------------------------------------------------------------------------------------------------%
for i = 1:24
    for j = 1:N
        node_E(j,i) = carbon(i).node_E(j);
        node_gen(j,i) = carbon(i).gen_node(j);
        node_CR(j,i) = carbon(i).node_CR(j);
        node_CU(j,i) = carbon(i).node_CU(j);
        node_CF(j,i)  = carbon(i).node_CF(j);
    end   
end
    title = {'节点ID','1','2','3','4','5','6','7','8','9','10','11','12','13','14','15','16','17','18','19','20','21','22','23','24'};
    node_E_temp = [node_data(:,1),node_E];                         node_gen_temp = [node_data(:,1),node_gen];
    node_CR_temp = [node_data(:,1),node_CR];                   node_CF_temp = [node_data(:,1),node_CF];
    node_CU_temp = [node_data(:,1),node_CU];  
    node_E_out = [title;num2cell(node_E_temp);];        node_gen_out = [title;num2cell(node_gen_temp)];
    node_CR_out = [title;num2cell(node_CR_temp)];     node_CF_out = [title;num2cell(node_CF_temp)];
    node_CU_out = [title;num2cell(node_CU_temp)]; 
    pathout     =   ['节点运行数据_new.xlsx'];
    sheetname_temp     =    ['节点碳势'];                    xlswrite(pathout,node_E_out,sheetname_temp);
    sheetname_temp     =    ['节点总发电量'];             xlswrite(pathout,node_gen_out,sheetname_temp);    
    sheetname_temp     =    ['节点直接碳排放量'];     xlswrite(pathout,node_CR_out,sheetname_temp);
    sheetname_temp     =    ['节点间接碳排放量'];     xlswrite(pathout,node_CU_out,sheetname_temp);    
    sheetname_temp     =    ['节点碳流率'];                xlswrite(pathout,node_CF_out,sheetname_temp);       
%------------------------------------------------------------------------------------------------------------------%
for i = 1:24
    for j = 1:P
         line_E(j,i) = carbon(i).line_E(j,1);
         line_CF(j,i) = carbon(i).line_CF(j);
         line_Closs(j,i) = carbon(i).line_Closs(j);
    end
end
    title = {'起始节点ID','终止节点ID','1','2','3','4','5','6','7','8','9','10','11','12','13','14','15','16','17','18','19','20','21','22','23','24'};
    line_E_temp = [line_data(:,2:3),line_E];                         line_CF_temp = [line_data(:,2:3),line_CF];
    line_Closs_temp = [line_data(:,2:3),line_Closs];
    line_E_out = [title;num2cell(line_E_temp)];       line_CF_out = [title;num2cell(line_CF_temp)];
    line_Closs_out = [title;num2cell(line_Closs_temp)];
    pathout     =   ['线路运行数据_new.xlsx'];
    sheetname_temp     =    ['线路碳流密度'];                    xlswrite(pathout,line_E_out,sheetname_temp);
    sheetname_temp     =    ['线路碳流量'];             xlswrite(pathout,line_CF_out,sheetname_temp);       
    sheetname_temp     =    ['网损碳排放量'];             xlswrite(pathout,line_Closs_out,sheetname_temp);      
%------------------------------------------------------------------------------------------------------------------%
for i = 1:24
    for j = 1:K
         gen_Cpai(j,i) = carbon(i).gen_Cpai(j);
    end
end      
    title = {'机组编号','机组所在节点ID','1','2','3','4','5','6','7','8','9','10','11','12','13','14','15','16','17','18','19','20','21','22','23','24'};
    gen_Cpai_temp = [gen_data(:,1:2),gen_Cpai];                      
    gen_Cpai_out = [title;num2cell(gen_Cpai_temp)];      
    pathout     =   ['机组运行数据_new.xlsx'];
    sheetname_temp     =    ['机组碳排放量'];                    xlswrite(pathout,gen_Cpai_out,sheetname_temp);
    %------------------------------------------------------------------------------------------------------------------%
  for i = 1:24
    for j = 1:Q
        area_load(j,i) = carbon(i).area_load(j);
        area_gen(j,i) = carbon(i).area_gen(j);
        area_E(j,i) = carbon(i).area_E(j);
        area_CR(j,i) = carbon(i).area_CR(j);
        area_CU(j,i) = carbon(i).area_CU(j);
        area_CF(j,i) = carbon(i).area_CF(j);
    end
end        
     title = {'区域编号','1','2','3','4','5','6','7','8','9','10','11','12','13','14','15','16','17','18','19','20','21','22','23','24'};
    area_load_temp = [region_data(:,1),area_load];              area_load_out = [title;num2cell(area_load_temp)];     
    area_gen_temp = [region_data(:,1),area_gen];                area_gen_out = [title;num2cell(area_gen_temp)];     
    area_E_temp = [region_data(:,1),area_E];                          area_E_out = [title;num2cell(area_E_temp)];     
    area_CR_temp = [region_data(:,1),area_CR];                    area_CR_out = [title;num2cell(area_CR_temp)];     
    area_CU_temp = [region_data(:,1),area_CU];                    area_CU_out = [title;num2cell(area_CU_temp)];     
    area_CF_temp = [region_data(:,1),area_CF];                     area_CF_out = [title;num2cell(area_CF_temp)];          
    pathout     =   ['区域运行数据_new.xlsx'];
    sheetname_temp     =    ['区域总负荷'];                        xlswrite(pathout,area_load_out,sheetname_temp);       
    sheetname_temp     =    ['区域总发电量'];                    xlswrite(pathout,area_gen_out,sheetname_temp);          
    sheetname_temp     =    ['区域碳势'];                            xlswrite(pathout,area_E_out,sheetname_temp);      
    sheetname_temp     =    ['区域直接碳排放量'];             xlswrite(pathout,area_CR_out,sheetname_temp);  
    sheetname_temp     =    ['区域间接碳排放量'];             xlswrite(pathout,area_CU_out,sheetname_temp);          
    sheetname_temp     =    ['区域碳流率'];                        xlswrite(pathout,area_CF_out,sheetname_temp);          
        
