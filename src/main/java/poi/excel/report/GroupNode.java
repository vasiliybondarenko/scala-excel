package poi.excel.report;

import java.util.List;

/**
 * Created with IntelliJ IDEA.
 * Author: shredinger
 * Date: 1/18/15
 * Time: 9:17 PM
 * Project: scala-excel
 */
public class GroupNode {
    private final String key;
    private final String value;
    private final List<GroupNode> children;
    private final List<Integer> leaves;

    public GroupNode(String key, String value, List<GroupNode> children, List<Integer> leaves) {
        this.key = key;
        this.value = value;
        this.children = children;
        this.leaves = leaves;
    }

    public String getKey() {
        return key;
    }

    public String getValue() {
        return value;
    }

    public List<GroupNode> getChildren() {
        return children;
    }

    public List<Integer> getLeaves() {
        return leaves;
    }
}
