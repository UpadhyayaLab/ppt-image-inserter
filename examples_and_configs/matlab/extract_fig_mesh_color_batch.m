function extract_fig_mesh_color_batch(manifestPath)
% EXTRACT_FIG_MESH_COLOR_BATCH  Run many .fig->.mat extractions in one session.
%
%   extract_fig_mesh_color_batch(MANIFESTPATH) reads a UTF-8 text file with one
%   job per line, tab-separated as "<figPath>\t<outMat>", and calls
%   EXTRACT_FIG_MESH_COLOR on each. A per-line try/catch keeps one bad figure
%   from aborting the batch. Status is printed as "OK <outMat>" /
%   "FAIL <figPath> : <msg>" so the caller can tell which extractions landed.
%
%   Running the whole batch in a single MATLAB -batch invocation avoids paying
%   MATLAB's ~10-15 s startup cost per figure.
    fid = fopen(manifestPath, 'r', 'n', 'UTF-8');
    if fid < 0
        error('extract_fig_mesh_color_batch:manifest', ...
              'cannot open manifest %s', manifestPath);
    end
    closer = onCleanup(@() fclose(fid));

    nOk = 0; nFail = 0;
    while true
        line = fgetl(fid);
        if ~ischar(line), break; end
        line = strtrim(line);
        if isempty(line), continue; end
        parts = strsplit(line, sprintf('\t'));
        if numel(parts) ~= 2
            fprintf('FAIL malformed-line : %s\n', line);
            nFail = nFail + 1;
            continue;
        end
        try
            extract_fig_mesh_color(parts{1}, parts{2});
            fprintf('OK %s\n', parts{2});
            nOk = nOk + 1;
        catch err
            fprintf('FAIL %s : %s\n', parts{1}, err.message);
            nFail = nFail + 1;
        end
    end
    fprintf('DONE ok=%d fail=%d\n', nOk, nFail);
end
