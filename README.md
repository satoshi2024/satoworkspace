<input type="checkbox" id="xxxItem" value="1" onclick="handleCheckboxClick()"
    <c:if test="${PAGE_VIEW.xxxItem == '1'}">checked="checked"</c:if>>
<label for="xxxItem">？？？</label>

<c:choose>
    <c:when test="${PAGE_VIEW.xxxItem == '1'}">
        <input type="hidden" id="xxxStatus" name="xxxStatus" value="1">
    </c:when>
    <c:otherwise>
        <input type="hidden" id="xxxStatus" name="xxxStatus" value="0">
    </c:otherwise>
</c:choose>