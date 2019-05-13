package group;

import com.alibaba.fastjson.JSON;
import group.GroupDO.GroupTestDO;

import java.util.ArrayList;
import java.util.List;
import java.util.Set;
import java.util.concurrent.ConcurrentHashMap;
import java.util.function.Function;
import java.util.function.Predicate;
import java.util.stream.Collectors;

/**
 * @author stone tiger
 * @Description: N个对象某一个字段重复，需要根据某个字段去重
 * @date 2019/5/13
 */
public class GroupBySegment {

    public static void main(String[] args) {
        List<GroupTestDO> list = new ArrayList<>();


        for (int i  = 0; i < 10 ;i ++){
            GroupTestDO groupTestDO1 = new GroupTestDO();
            groupTestDO1.setId(i);
            groupTestDO1.setName("name"+i);
            GroupTestDO groupTestDO2 = new GroupTestDO();
            groupTestDO2.setId(i);
            groupTestDO2.setName("name"+i);
            list.add(groupTestDO1);
            list.add(groupTestDO2);
        }
        list =list.stream().filter(distinctByKey(GroupTestDO::getId)).collect(Collectors.toList());
        System.out.println(JSON.toJSONString(list));

    }

    public static <T> Predicate<T> distinctByKey(Function<? super T, ?> keyExtractor) {
        Set<Object> seen = ConcurrentHashMap.newKeySet();
        return t -> seen.add(keyExtractor.apply(t));
    }
}
