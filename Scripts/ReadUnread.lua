function Initialize()
    file = SELF:GetOption('File')
end

function Update()
    local f = io.open(file, "r")
    if not f then return 0 end

    for line in f:lines() do
        local count = string.match(line, "UnreadCount=(%d+)")
        if count then
            f:close()
            return tonumber(count)
        end
    end

    f:close()
    return 0
end